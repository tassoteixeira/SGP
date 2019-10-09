VERSION 5.00
Begin VB.Form movimento_cartao_credito 
   Caption         =   "Movimentação do Cartão de Crédito"
   ClientHeight    =   4365
   ClientLeft      =   2325
   ClientTop       =   3105
   ClientWidth     =   6975
   Icon            =   "movimento_cartao_credito.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "movimento_cartao_credito.frx":030A
   ScaleHeight     =   4365
   ScaleWidth      =   6975
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "movimento_cartao_credito.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Cria um novo registro."
      Top             =   3420
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "movimento_cartao_credito.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Altera o registro atual."
      Top             =   3420
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "movimento_cartao_credito.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Exclui o registro atual."
      Top             =   3420
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "movimento_cartao_credito.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   3420
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "movimento_cartao_credito.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3420
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6735
      Begin VB.TextBox txtValor 
         Height          =   285
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   17
         Top             =   2100
         Width           =   1095
      End
      Begin VB.TextBox txtDataVencimento 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   15
         Top             =   1740
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
      Begin VB.TextBox txtNSU 
         Height          =   285
         Left            =   5520
         MaxLength       =   10
         TabIndex        =   25
         Top             =   2820
         Width           =   1095
      End
      Begin VB.TextBox txtAutorizacao 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   23
         Top             =   2820
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
      Begin VB.ComboBox cboCartao 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1320
         Width           =   4515
      End
      Begin VB.TextBox txt_cartao 
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   12
         Top             =   1320
         Width           =   375
      End
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_numero_cartao 
         Height          =   285
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   19
         Top             =   2460
         Width           =   735
      End
      Begin VB.ComboBox cbo_tipo_movimento 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   2475
      End
      Begin VB.Label Label11 
         Caption         =   "&NSU/DOC"
         Height          =   255
         Left            =   4500
         TabIndex        =   24
         Top             =   2850
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "&N. da Autorização"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2820
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "I&lha"
         Height          =   255
         Left            =   4920
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lbl_numero_lancamento 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "C&artão"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Total"
         Height          =   255
         Left            =   4860
         TabIndex        =   20
         Top             =   2460
         Width           =   555
      End
      Begin VB.Label lbl_total 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   21
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "&Data do Vencimento"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "N. do lancamento"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "&Número do Cartão"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   2460
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "&Tipo de Movimento"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &Emissão"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "&Período"
         Height          =   255
         Left            =   4920
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "&Valor do Cartão"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   1455
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4680
      TabIndex        =   33
      Top             =   3300
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "movimento_cartao_credito.frx":7472
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "movimento_cartao_credito.frx":896C
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "movimento_cartao_credito.frx":9E66
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "movimento_cartao_credito.frx":B2D8
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5160
      Picture         =   "movimento_cartao_credito.frx":C85A
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Confirma o registro atual."
      Top             =   3420
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6060
      Picture         =   "movimento_cartao_credito.frx":DE64
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Cancela o registro atual."
      Top             =   3420
      Width           =   795
   End
End
Attribute VB_Name = "movimento_cartao_credito"
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
Dim lOrdem As Integer
Dim lTotal As Currency
Dim lNumeroMovimentoCaixa As Long
Dim lGravados As Integer
Dim lQtdPeriodo As Integer
Dim lCartaoAnterior As Integer
Dim lValorAnterior As Currency
Dim lCaixaIndividual As Boolean
Dim lNomeAnterior As String

Dim lCxData As Date
Dim lCxPeriodo As String
Dim lCxOrdem As Integer
Dim lCxTipoMov As Integer
Dim lCxTipoMovimento As Integer
Dim lCxIlha As Integer
Dim lCxDataDigitacao As Date
Dim lCxHoraDigitacao As Date
Dim lCxCodigoLancamentoPadrao As Integer
Dim lCxCodigoFuncionario As Integer
Dim lCxCodigoUsuario As Integer
Dim lCxValor As Currency
Dim lCxNumeroNFCe As Long
Dim lLancamentoDePOS As Boolean
Dim lLancamentoDeBaixaDuplicata As Boolean
Dim lNumeroDocumentoDuplicata As String

Dim lOrigem As String
Dim lCxCodigoCartaoPreDefinido As Integer


Private IntegracaoCaixa As New cIntegracaoCaixa
Private CartaoCredito As New cCartaoCredito
Private Configuracao As New cConfiguracao
Private ConfiguracaoDiversa As New cConfiguracaoDiversa
Private MovimentoCaixaPista As New cMovimentoCaixaPista
Private MovCartaoCredito As New cMovimentoCartaoCredito
Private TaxaAdmCartaoCredito As New cTaxaAdmCartaoCredito

Private Sub AtualizaConstantes()
    If Configuracao.LocalizarCodigo(g_empresa) Then
        lQtdPeriodo = Configuracao.QuantidadePeriodos
    Else
        lQtdPeriodo = 1
    End If
    lLancamentoDePOS = False
    lLancamentoDeBaixaDuplicata = False
End Sub
Private Sub AtualizaTabela()
    MovCartaoCredito.Empresa = g_empresa
    MovCartaoCredito.DataEmissao = CDate(txtDataEmissao.Text)
    MovCartaoCredito.Periodo = Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
    MovCartaoCredito.TipoMovimento = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
    MovCartaoCredito.NumeroLancamento = Val(lbl_numero_lancamento.Caption)
    MovCartaoCredito.CodigoCartao = Val(cboCartao.ItemData(cboCartao.ListIndex))
    MovCartaoCredito.DataVencimento = CDate(txtDataVencimento.Text)
    MovCartaoCredito.Valor = fValidaValor(txtValor.Text)
    MovCartaoCredito.NumeroCartao = txt_numero_cartao.Text
    MovCartaoCredito.Nome = ""
    If lLancamentoDePOS = True Then
        MovCartaoCredito.Nome = "NFCe-POS: " & lCxNumeroNFCe
    ElseIf lLancamentoDeBaixaDuplicata = True Then
        MovCartaoCredito.Nome = "DUPL-POS: " & lNumeroDocumentoDuplicata
    End If
    If lOpcao = 2 Then
        MovCartaoCredito.Nome = lNomeAnterior
    End If
    MovCartaoCredito.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
    MovCartaoCredito.TaxaAdministrativa = TaxaAdmCartaoCredito.TaxaCusto
    MovCartaoCredito.NumeroIlha = Val(cboIlha.Text)
    MovCartaoCredito.Autorizacao = txtAutorizacao.Text
    MovCartaoCredito.NSU = txtNSU.Text
    
    If lCxCodigoFuncionario = 0 Then
        MovCartaoCredito.CodigoFuncionario = ObtenhaCodigoFuncionarioDoUsuario
    Else
        MovCartaoCredito.CodigoFuncionario = lCxCodigoFuncionario
    End If
    
End Sub

Private Function ObtenhaCodigoFuncionarioDoUsuario() As Integer

    ObtenhaCodigoFuncionarioDoUsuario = 0
    Dim xFuncionario As New CadastroDLL.cFuncionario
    
    If (xFuncionario.LocalizarFuncionarioDoUsuario(g_usuario, g_empresa)) Then
        ObtenhaCodigoFuncionarioDoUsuario = xFuncionario.Codigo
    End If
    
    Exit Function

End Function

Private Sub PreencheCboTipoMovimento()
    cbo_tipo_movimento.Clear
    cbo_tipo_movimento.AddItem "1 - Caixa de Combustíveis"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 1
    cbo_tipo_movimento.AddItem "2 - Caixa de Óleos/Diversos"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 2
    cbo_tipo_movimento.AddItem "3 Caixa da Borr./Lavador"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 3
    cbo_tipo_movimento.AddItem "4 - Cartão POS/Sem Caixa"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 4
End Sub
Private Sub AtualizaTela()
    Dim i As Integer
    
    lData = MovCartaoCredito.DataEmissao
    lPeriodo = MovCartaoCredito.Periodo
    lIlha = MovCartaoCredito.NumeroIlha
    lOrdem = MovCartaoCredito.NumeroLancamento
    lNumeroMovimentoCaixa = MovCartaoCredito.NumeroMovimentoCaixa
    lCartaoAnterior = MovCartaoCredito.CodigoCartao
    lValorAnterior = MovCartaoCredito.Valor
    lNomeAnterior = MovCartaoCredito.Nome

    txtDataEmissao.Text = Format(MovCartaoCredito.DataEmissao, "dd/mm/yyyy")
    cbo_periodo.ListIndex = Val(MovCartaoCredito.Periodo) - 1
    cbo_tipo_movimento.ListIndex = MovCartaoCredito.TipoMovimento - 1
    cboIlha.ListIndex = Val(MovCartaoCredito.NumeroIlha) - 1
    lbl_numero_lancamento.Caption = MovCartaoCredito.NumeroLancamento
    txt_cartao.Text = MovCartaoCredito.CodigoCartao
    If TaxaAdmCartaoCredito.LocalizarCodigo(g_empresa, MovCartaoCredito.CodigoCartao) Then
    Else
        TaxaAdmCartaoCredito.TaxaCusto = CartaoCredito.TaxaCusto
        MsgBox "Taxa de Adm de Cartão de crédito não cadastrada.", vbInformation, "Erro de Integridade!"
    End If
    
    Call SelecionaCartaoNaCombo(MovCartaoCredito.CodigoCartao)
'    cboCartao.ListIndex = -1
'    For i = 0 To cboCartao.ListCount - 1
'        If cboCartao.ItemData(i) = MovCartaoCredito.CodigoCartao Then
'            cboCartao.ListIndex = i
'            Exit For
'        End If
'    Next
    txtDataVencimento.Text = Format(MovCartaoCredito.DataVencimento, "dd/mm/yyyy")
    txtValor.Text = Format(MovCartaoCredito.Valor, "###,##0.00")
    txt_numero_cartao.Text = MovCartaoCredito.NumeroCartao
    txtAutorizacao.Text = MovCartaoCredito.Autorizacao
    txtNSU.Text = MovCartaoCredito.NSU
    lbl_total.Caption = Format(MovCartaoCredito.TotalPeriodo(g_empresa, lData, lPeriodo, MovCartaoCredito.TipoMovimento, MovCartaoCredito.CodigoCartao), "###,##0.00")
    lNumeroMovimentoCaixa = MovCartaoCredito.NumeroMovimentoCaixa
    frm_dados.Enabled = False
    VerificaLiberacaoDigitacao
End Sub
Private Sub AtualizaTelaDadosCartaoSelecionado()

    If cboCartao.ListIndex <> -1 And lOpcao > 0 Then
        txt_cartao.Text = cboCartao.ItemData(cboCartao.ListIndex)
        lbl_numero_lancamento.Caption = MovCartaoCredito.ProximoRegistro(g_empresa, CDate(txtDataEmissao.Text))
        If CartaoCredito.LocalizarCodigo(Val(cboCartao.ItemData(cboCartao.ListIndex))) Then
            If TaxaAdmCartaoCredito.LocalizarCodigo(g_empresa, CartaoCredito.Codigo) Then
            Else
                TaxaAdmCartaoCredito.TaxaCusto = CartaoCredito.TaxaCusto
                MsgBox "Taxa de Adm de Cartão de crédito não cadastrada.", vbInformation, "Erro de Integridade!"
            End If
            If IsDate(txtDataEmissao.Text) Then
                txtDataVencimento.Text = CDate(txtDataEmissao.Text) + CartaoCredito.DiasPrazo
            End If
        Else
            MsgBox "Cartão de crédito não cadastrado.", vbInformation, "Erro de Integridade!"
        End If
        txt_cartao_LostFocus
        If txtValor.Enabled = True Then
            txtValor.SetFocus
        Else
            If lOpcao = 1 And txt_numero_cartao.Text = "" Then
                txt_numero_cartao.Text = 1
            End If
            txtAutorizacao.SetFocus
        End If
    End If
End Sub
Private Sub SelecionaCartaoNaCombo(ByVal pCodigoCartao As Integer)
    Dim i As Integer
    cboCartao.ListIndex = -1
    For i = 0 To cboCartao.ListCount - 1
        If cboCartao.ItemData(i) = pCodigoCartao Then
            cboCartao.ListIndex = i
            Exit For
        End If
    Next
End Sub
Private Sub SelecionaTipoMovimentoNaCombo(ByVal pTipoMovimento As Integer)
    Dim i As Integer
    cbo_tipo_movimento.ListIndex = -1
    For i = 0 To cbo_tipo_movimento.ListCount - 1
        If cbo_tipo_movimento.ItemData(i) = pTipoMovimento Then
            cbo_tipo_movimento.ListIndex = i
            Exit For
        End If
    Next
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
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    FinalizaProcessoCaixa
    Set CartaoCredito = Nothing
    Set Configuracao = Nothing
    Set ConfiguracaoDiversa = Nothing
    Set IntegracaoCaixa = Nothing
    Set MovimentoCaixaPista = Nothing
    Set MovCartaoCredito = Nothing
    Set TaxaAdmCartaoCredito = Nothing
End Sub
Private Sub Inclui()
    lOpcao = 1
    lNomeAnterior = ""
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Function IncluiMovimentoCaixa() As Boolean
    Dim xComplemento As String
    IncluiMovimentoCaixa = False
    lNumeroMovimentoCaixa = 0
    
    If IntegracaoCaixa.LocalizarNome(g_empresa, "CARTAO " & cboCartao.Text) Then
        xComplemento = "P/ " & txtDataVencimento.Text & " TM:" & cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex) & " P:" & Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
        
        
        
        MovimentoCaixaPista.Empresa = g_empresa
        MovimentoCaixaPista.Data = CDate(txtDataEmissao.Text)
        MovimentoCaixaPista.NumeroMovimento = 1
        MovimentoCaixaPista.Valor = fValidaValor(txtValor.Text)
        MovimentoCaixaPista.NumeroDocumento = txt_numero_cartao.Text
        MovimentoCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
        MovimentoCaixaPista.Complemento = Mid(xComplemento, 1, 40)
        MovimentoCaixaPista.NumeroContaDebito = IntegracaoCaixa.ContaDebito
        MovimentoCaixaPista.NumeroContaCredito = IntegracaoCaixa.ContaCredito
        MovimentoCaixaPista.TipoMovimento = lCxTipoMovimento
        MovimentoCaixaPista.CodigoUsuario = lCxCodigoUsuario
        MovimentoCaixaPista.Periodo = Val(cbo_periodo.Text)
        MovimentoCaixaPista.NumeroIlha = Val(cboIlha.Text)
        MovimentoCaixaPista.DadosInterno = "CAR" & Format(Val(txt_cartao.Text), "00") & "|@|" & lbl_numero_lancamento.Caption & "|@|"
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
        MsgBox "Não existe a integração=" & "CARTAO " & cboCartao.Text & ".", vbInformation, "Registro Inexistente"
    End If
End Function
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
Private Sub cbo_tipo_movimento_LostFocus()
'    If lOpcao = 1 Then
'        lbl_numero_lancamento.Caption = "1"
'    End If
End Sub
Private Sub cboCartao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtValor.Enabled = True Then
            txtValor.SetFocus
        Else
            If lOpcao = 1 And txt_numero_cartao.Text = "" Then
                txt_numero_cartao.Text = 1
            End If
            txtAutorizacao.SetFocus
        End If
    End If
End Sub
Private Sub cboCartao_LostFocus()
    If cboCartao.ListIndex <> -1 And lOpcao > 0 Then
        txt_cartao.Text = cboCartao.ItemData(cboCartao.ListIndex)
        'lbl_numero_lancamento.Caption = MovCartaoCredito.ProximoRegistro(g_empresa, CDate(txtDataEmissao.Text), cbo_periodo.Text)
        lbl_numero_lancamento.Caption = MovCartaoCredito.ProximoRegistro(g_empresa, CDate(txtDataEmissao.Text))
        If CartaoCredito.LocalizarCodigo(Val(cboCartao.ItemData(cboCartao.ListIndex))) Then
            If TaxaAdmCartaoCredito.LocalizarCodigo(g_empresa, CartaoCredito.Codigo) Then
            Else
                TaxaAdmCartaoCredito.TaxaCusto = CartaoCredito.TaxaCusto
                MsgBox "Taxa de Adm de Cartão de crédito não cadastrada.", vbInformation, "Erro de Integridade!"
            End If
            If IsDate(txtDataEmissao.Text) Then
                txtDataVencimento.Text = CDate(txtDataEmissao.Text) + CartaoCredito.DiasPrazo
            End If
        Else
            MsgBox "Cartão de crédito não cadastrado.", vbInformation, "Erro de Integridade!"
        End If
        txt_cartao_LostFocus
        If txtValor.Enabled = True Then
            txtValor.SetFocus
        Else
            If lOpcao = 1 And txt_numero_cartao.Text = "" Then
                txt_numero_cartao.Text = 1
            End If
            txtAutorizacao.SetFocus
        End If
    End If
End Sub
Private Sub cboIlha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_cartao.SetFocus
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
    If MovCartaoCredito.LocalizarAnterior Then
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
    If MovCartaoCredito.LocalizarCodigo(g_empresa, lData, lPeriodo, lOrdem) Then
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
    If lGravados = 0 Then
        txtDataEmissao.Text = ""
        cbo_periodo.ListIndex = -1
        cbo_tipo_movimento.ListIndex = -1
        cboIlha.ListIndex = -1
        lbl_numero_lancamento.Caption = ""
        txt_cartao.Text = ""
        cboCartao.ListIndex = -1
        txtDataVencimento.Text = ""
    End If
    txtValor.Text = ""
    txt_numero_cartao.Text = ""
    txtAutorizacao.Text = ""
    txtNSU.Text = ""
End Sub
Private Sub cmd_excluir_Click()
    Call GravaAuditoria(1, Me.name, 4, "")
    If IsDate(txtDataVencimento.Text) Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            Call GravaAuditoria(1, Me.name, 10, "Dt:" & txtDataVencimento.Text & " Per:" & cbo_periodo.Text & " N.Mov:" & lbl_numero_lancamento.Caption & " Vlr:" & txtValor.Text & " Cod.Cart:" & txt_cartao.Text)
            Call ExcluiMovimentoCaixa
            lOpcao = 3
            If MovCartaoCredito.Excluir(g_empresa, CDate(txtDataEmissao.Text), cbo_periodo.Text, Val(lbl_numero_lancamento.Caption)) Then
                'If Not MovimentoCaixaPista.Excluir(g_empresa, lData, lNumeroMovimentoCaixa) Then
                '    MsgBox "Não foi excluído o movimento do caixa!", vbInformation, "Erro de Integridade."
                'End If
            End If
            If MovCartaoCredito.LocalizarUltimo(g_empresa) Then
                AtualizaTela
            Else
                DesativaBotoes
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
                cmd_novo.SetFocus
            End If
            lOpcao = 0
        End If
    End If
    If lCxPeriodo > 0 Then
        cmd_sair_Click
        Exit Sub
    End If
End Sub
Private Sub cmd_novo_Click()
    'Exit Sub
    'Call zzLancaCartaoPelaComposicao
    '
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
            txt_cartao.SetFocus
            If lLancamentoDePOS = True Then
                txtValor.Text = Format(lCxValor, "###,##0.00")
                txtValor.Enabled = False
                'txtNSU.Enabled = False
            ElseIf lLancamentoDeBaixaDuplicata = True Then
                txtValor.Text = Format(lCxValor, "###,##0.00")
                txtValor.Enabled = False
            End If
            Exit Sub
        End If
        If BuscaProximoCaixa Then
            txt_cartao.SetFocus
        Else
            txtDataEmissao.SetFocus
        End If
    Else
        'lbl_numero_lancamento.Caption = MovCartaoCredito.ProximoRegistro(g_empresa, CDate(txtDataEmissao.Text), cbo_periodo.Text)
        lbl_numero_lancamento.Caption = MovCartaoCredito.ProximoRegistro(g_empresa, CDate(txtDataEmissao.Text))
        txtValor.SetFocus
    End If
End Sub
Function BuscaProximoCaixa() As Boolean
    Dim x_periodo As String
    BuscaProximoCaixa = False
    If MovCartaoCredito.LocalizarUltimo(g_empresa) Then
        txtDataEmissao.Text = MovCartaoCredito.DataEmissao
        x_periodo = MovCartaoCredito.Periodo
        If MovCartaoCredito.Periodo >= lQtdPeriodo Then
            txtDataEmissao.Text = MovCartaoCredito.DataEmissao + 1
            x_periodo = 0
        End If
        cbo_periodo.ListIndex = x_periodo
        cbo_tipo_movimento.ListIndex = 0
        cboIlha.ListIndex = 0
        BuscaProximoCaixa = True
    Else
        txtDataEmissao.Text = g_data_def - 1
        cbo_periodo.ListIndex = 0
        cbo_tipo_movimento.ListIndex = 0
        cboIlha.ListIndex = 0
    End If
End Function
Private Sub cmd_novo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 16 Then
        MsgBox "PROCESSAMENTO"
        Call ProcessaCartaoCredito
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        If VerificaLiberacaoDigitacao2 Then
            AtivaBotoes
            If lOpcao = 1 Then
                
                If Not lLancamentoDeBaixaDuplicata Then
                    If Not IncluiMovimentoCaixa Then
                        MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
                    End If
                End If
                
                AtualizaTabela
                
                Call GravaAuditoria(1, Me.name, 10, "Dt:" & txtDataVencimento.Text & " Per:" & cbo_periodo.Text & " N.Mov:" & lbl_numero_lancamento.Caption & " Vlr:" & txtValor.Text & " Cod.Cart:" & txt_cartao.Text)
'ver aqui
                Dim xNumeroLancamentoAtual As Integer
                xNumeroLancamentoAtual = MovCartaoCredito.NumeroLancamento
                
                MovCartaoCredito.NumeroLancamento = MovCartaoCredito.ProximoRegistro(g_empresa, CDate(txtDataEmissao.Text))
                lbl_numero_lancamento.Caption = MovCartaoCredito.NumeroLancamento
                
                If MovCartaoCredito.Incluir Then
                    
                    If xNumeroLancamentoAtual <> MovCartaoCredito.NumeroLancamento Then
                        Dim xDadosInternoAtual As String
                        xDadosInternoAtual = MovimentoCaixaPista.DadosInterno
                        MovimentoCaixaPista.DadosInterno = "CAR" & Format(Val(txt_cartao.Text), "00") & "|@|" & MovCartaoCredito.NumeroLancamento & "|@|"
                        If Not MovimentoCaixaPista.DefineDadosInternoMovimentoCaixaDoCartao(g_empresa, CDate(txtDataEmissao.Text), CInt(cbo_periodo.Text), MovCartaoCredito.NumeroMovimentoCaixa, xDadosInternoAtual) Then
                            MsgBox "Não foi possível alterar dados integração com o Caixa!", vbInformation, "Erro de Integridade."
                            Call CriaLogSGP("[cmd_ok_Click] - MovimentoCaixaPista.DefineDadosInternoMovimentoCaixaDoCartao", "Não foi possível alterar dados integração com o Caixa", "")
                        End If
                    End If
                    
                    lData = MovCartaoCredito.DataEmissao
                    lPeriodo = MovCartaoCredito.Periodo
                    lOrdem = MovCartaoCredito.NumeroLancamento
                    'If Not IncluiMovimentoCaixa Then
                    '    MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
                    'End If
                    lGravados = 1
                End If
            ElseIf lOpcao = 2 Then
                If lCxPeriodo > 0 Then
                    Call ExcluiMovimentoCaixa
                    If Not IncluiMovimentoCaixa Then
                        MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
                    End If
                End If
                AtualizaTabela
                Call GravaAuditoria(1, Me.name, 10, "De: Dt:" & lData & " Per:" & lPeriodo & " N.Mov:" & lOrdem & " Vlr:" & lValorAnterior & " Cod.Cart:" & lCartaoAnterior)
                Call GravaAuditoria(1, Me.name, 10, "Para: Dt:" & txtDataVencimento.Text & " Per:" & cbo_periodo.Text & " N.Mov:" & lbl_numero_lancamento.Caption & " Vlr:" & txtValor.Text & " Cod.Cart:" & txt_cartao.Text)
                If MovCartaoCredito.Alterar(g_empresa, lData, lPeriodo, lOrdem) Then
                    lData = MovCartaoCredito.DataEmissao
                    lPeriodo = MovCartaoCredito.Periodo
                    lOrdem = MovCartaoCredito.NumeroLancamento
                End If
            End If
'ver aqui
            If MovCartaoCredito.LocalizarCodigo(g_empresa, lData, lPeriodo, lOrdem) Then
                AtualizaTela
            Else
                LimpaTela
            End If
            If lOpcao = 1 Then
                lOpcao = 0
                If lLancamentoDePOS = True And lOrigem = "Movimento_NFCe_Auto" Then
                    g_string = "Retorno-Movimento_NFCe_Auto|@|" & cboCartao.ItemData(cboCartao.ListIndex) & "|@|" & cboCartao.Text & "|@|" & txtValor.Text & "|@|" & txtAutorizacao.Text & "|@|"
                    cmd_sair_Click
                    Exit Sub
                ElseIf lLancamentoDePOS = True And lOrigem = "Movimento_Nfce_Conveniencia" Then
                    g_string = "Retorno-Movimento_Nfce_Conveniencia|@|" & cboCartao.ItemData(cboCartao.ListIndex) & "|@|" & cboCartao.Text & "|@|" & txtValor.Text & "|@|" & txtAutorizacao.Text & "|@|"
                    cmd_sair_Click
                    Exit Sub
                ElseIf lLancamentoDeBaixaDuplicata = True Then
                    'Format(CartaoCredito.Codigo, "00") & "|@|" & lNumeroLancamentoCartao & "|@|
                    g_string = "Retorno-BaixaDuplicataReceber|@|" & cboCartao.ItemData(cboCartao.ListIndex) & "|@|" & lbl_numero_lancamento.Caption & "|@|"
                    cmd_sair_Click
                    Exit Sub
                Else
                    cmd_novo_Click
                End If
            Else
                If lCxPeriodo > 0 Then
                    cmd_sair_Click
                    Exit Sub
                End If
                lOpcao = 0
                cmd_novo.SetFocus
            End If
        End If
    End If
    Exit Sub
FileError:
    'ErroArquivo tbl_movimento_cartao_credito.Name, "Movimento de Cartaoo"
    MsgBox Err & " - " & Error
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(txtDataEmissao.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        txtDataEmissao.SetFocus
    ElseIf cbo_periodo.ListIndex = -1 Then
        MsgBox "Selecione o período.", vbInformation, "Atenção!"
        cbo_periodo.SetFocus
    ElseIf cbo_tipo_movimento.ListIndex = -1 Then
        MsgBox "Selecione o tipo de movimento.", vbInformation, "Atenção!"
        cbo_tipo_movimento.SetFocus
    ElseIf cboIlha.ListIndex = -1 Then
        MsgBox "Selecione uma ilha.", vbInformation, "Atenção!"
        cboIlha.SetFocus
    ElseIf Not Val(lbl_numero_lancamento.Caption) > 0 Then
        MsgBox "Informe o número do lançamento.", vbInformation, "Atenção!"
        cbo_tipo_movimento.SetFocus
    ElseIf cboCartao.ListIndex = -1 Then
        MsgBox "Selecione o cartão de crédito.", vbInformation, "Atenção!"
        cboCartao.SetFocus
    ElseIf Not IsDate(txtDataVencimento.Text) Then
        MsgBox "Informe a data de vencimento.", vbInformation, "Atenção!"
        txtDataVencimento.SetFocus
    ElseIf Not fValidaValor(txtValor.Text) > 0 Then
        MsgBox "Informe o valor do cartão.", vbInformation, "Atenção!"
        txtValor.SetFocus
    ElseIf Not Val(txt_numero_cartao.Text) > 0 Then
        MsgBox "Informe o número do cartão.", vbInformation, "Atenção!"
        txt_numero_cartao.SetFocus
    ElseIf lLancamentoDePOS = True And txtAutorizacao.Text = "" Then
        MsgBox "Informe o número da autorização.", vbInformation, "Atenção!"
        txtAutorizacao.SetFocus
    ElseIf Trim(txtNSU.Text) = Empty Then
        MsgBox "Informe Doc/NSU da transação.", vbInformation, "Atenção!"
        txtNSU.SetFocus
    ElseIf IsNumeric(txtNSU.Text) And Val(txtNSU.Text) <= 0 Then
        MsgBox "Doc/NSU informado é inválido.", vbInformation, "Atenção!"
        txtNSU.SetFocus
    ElseIf Len(txtNSU.Text) < 4 Then
        MsgBox "Doc/NSU informado é inválido.", vbInformation, "Atenção!"
        txtNSU.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub VerificaEExcluiMovExistentes()
    If MovCartaoCredito.LocalizarDataPeriodoNome(g_empresa, lCxData, lCxPeriodo, lCxTipoMovimento, "NFCe-POS: " & lCxNumeroNFCe) = True Then
        If MovCartaoCredito.Excluir(g_empresa, lCxData, lCxPeriodo, MovCartaoCredito.NumeroLancamento) = True Then
            If MovimentoCaixaPista.Excluir(g_empresa, lCxData, MovCartaoCredito.NumeroMovimentoCaixa) = True Then
            Else
                MsgBox "Não foi possível excluir o movimento de caixa!", vbCritical + vbOKOnly, "Erro de Integridade!"
            End If
        Else
            MsgBox "Não foi possível excluir o movimento de cartão!", vbCritical + vbOKOnly, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub VerificaLiberacaoDigitacao()
    Dim x_flag As Boolean
    x_flag = True
    If g_nivel_acesso > 4 Then
        If MovCartaoCredito.Empresa < g_cfg_empresa_i Or MovCartaoCredito.Empresa > g_cfg_empresa_f Then
            x_flag = False
        ElseIf MovCartaoCredito.DataEmissao < g_cfg_data_i Or MovCartaoCredito.DataEmissao > g_cfg_data_f Then
            x_flag = False
        ElseIf MovCartaoCredito.Periodo < g_cfg_periodo_i Or MovCartaoCredito.Periodo > g_cfg_periodo_f Then
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
    ElseIf cbo_periodo.Text < g_cfg_periodo_i Or cbo_periodo.Text > g_cfg_periodo_f Then
        MsgBox "O período deve estar entre " & g_cfg_periodo_i & " ao " & g_cfg_periodo_f & ".", vbInformation, "Digitação Não Autorizada!"
        cbo_periodo.SetFocus
    Else
        VerificaLiberacaoDigitacao2 = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    Call GravaAuditoria(1, Me.name, 5, "")
    consulta_movimento_cartao.Show 1
    If Len(g_string) > 0 Then
        lData = RetiraGString(1)
        lPeriodo = RetiraGString(2)
        lOrdem = RetiraGString(3)
        g_string = ""
        If MovCartaoCredito.LocalizarCodigo(g_empresa, lData, lPeriodo, lOrdem) Then
            AtualizaTela
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    Call GravaAuditoria(1, Me.name, 12, "")
    If MovCartaoCredito.LocalizarPrimeiro Then
        AtualizaTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    Call GravaAuditoria(1, Me.name, 14, "")
    If MovCartaoCredito.LocalizarProximo Then
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
    If MovCartaoCredito.LocalizarUltimo(g_empresa) Then
        AtualizaTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub AjustaCaixaPista(ByVal pOrigem As String)
    Dim xString As String
    Dim xOperacao As String
    
    xString = g_string
    xOperacao = RetiraString(2, xString)
    g_string = ""
    lCxValor = 0
'g_string = "BaixaDuplicataReceber|@|Incluir|@|" & msk_data_pagamento.Text & "|@|" & cbo_periodo.Text & "|@|4|@|1|@|4|@|0|@|" & g_usuario & "|@|" & txtValorCartao.Text & "|@|" & DuplicataReceber.NumeroDocumento & "|@|"
    lCxData = CDate(RetiraString(3, xString))
    lCxPeriodo = RetiraString(4, xString)
    lCxTipoMovimento = Val(RetiraString(5, xString))
    
    lCxIlha = Val(RetiraString(6, xString))
    lCxTipoMov = Val(RetiraString(7, xString))
    lCxCodigoLancamentoPadrao = Val(RetiraString(8, xString))
    

    If pOrigem = "CaixaPista" And lCaixaIndividual Then
        lCxCodigoFuncionario = Val(RetiraString(11, xString))
    ElseIf pOrigem = "Movimento_NFCe_Auto" Or pOrigem = "Movimento_Nfce_Conveniencia" Then
        lCxCodigoFuncionario = Val(RetiraString(12, xString))
        If pOrigem = "Movimento_NFCe_Auto" Then
            lCxCodigoCartaoPreDefinido = Val(RetiraString(13, xString))
        End If
    ElseIf pOrigem = "BaixaDuplicataReceber" Then
        lCxCodigoFuncionario = 0
        lNumeroDocumentoDuplicata = ""
    End If

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
        If lLancamentoDePOS = True Then
            lCxValor = fValidaValor(RetiraString(10, xString))
            lCxNumeroNFCe = CLng(RetiraString(11, xString))
            VerificaEExcluiMovExistentes
        ElseIf lLancamentoDeBaixaDuplicata = True Then
            lCxValor = fValidaValor(RetiraString(10, xString))
            lCxNumeroNFCe = 0
            lNumeroDocumentoDuplicata = RetiraString(11, xString)
            lNumeroMovimentoCaixa = 0
        End If
        cmd_novo_Click
    ElseIf xOperacao = "Alterar" Then
        lCxOrdem = CLng(RetiraString(9, xString))
        lCxCodigoUsuario = Val(RetiraString(10, xString))
        If MovCartaoCredito.LocalizarCodigo(g_empresa, lCxData, lCxPeriodo, lCxOrdem) Then
            AtualizaTela
        End If
        cmd_alterar_Click
    ElseIf xOperacao = "Excluir" Then
        lCxOrdem = CLng(RetiraString(9, xString))
        lCxCodigoUsuario = Val(RetiraString(10, xString))
        If MovCartaoCredito.LocalizarCodigo(g_empresa, lCxData, lCxPeriodo, lCxOrdem) Then
            AtualizaTela
        Else
            g_string = "EXCLUSAO ESPECIAL|@|"
        End If
        cmd_excluir_Click
    End If
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
Private Sub Form_Activate()
    'Dim xOrigem As String
    lOrigem = ""
    If g_empresa <> lEmpresa Then
        lFlagMovimento = 0
    End If
    If lFlagMovimento = 0 Then
        AtualizaConstantes
        lOpcao = 0
        lEmpresa = g_empresa
        lGravados = 0
        DesativaBotoes
        'xOrigem = RetiraGString(1)
        lOrigem = RetiraGString(1)
        If lOrigem = "CaixaPista" Then
            'AjustaCaixaPista (xOrigem)
            AjustaCaixaPista (lOrigem)
        ElseIf lOrigem = "Movimento_NFCe_Auto" Or lOrigem = "Movimento_Nfce_Conveniencia" Then
            lLancamentoDePOS = True
            'AjustaCaixaPista (xOrigem)
            AjustaCaixaPista (lOrigem)
            If lCxCodigoCartaoPreDefinido > 0 Then
                Call SelecionaCartaoNaCombo(lCxCodigoCartaoPreDefinido)
                Call AtualizaTelaDadosCartaoSelecionado
                txt_cartao.Enabled = False
                cboCartao.Enabled = False
            End If
        ElseIf lOrigem = "BaixaDuplicataReceber" Then
            lLancamentoDeBaixaDuplicata = True
            AjustaCaixaPista (lOrigem)
        Else
            lCxPeriodo = 0
            If MovCartaoCredito.LocalizarUltimo(g_empresa) Then
                AtualizaTela
                AtivaBotoes
            Else
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
            End If
            If cmd_novo.Enabled = True Then
                cmd_novo.SetFocus
            End If
        End If
    Else
        lFlagMovimento = 0
    End If
    Screen.MousePointer = 1
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
    Call GravaAuditoria(1, Me.name, 1, "")
    CentraForm Me
    lCaixaIndividual = False
    If ConfiguracaoDiversa.LocalizarCodigo(g_empresa, "CAIXA DE PISTA INDIVIDUAL") Then
        lCaixaIndividual = ConfiguracaoDiversa.Verdadeiro
    End If
    PreencheCboIlha
    PreencheCboPeriodo
    PreencheCboTipoMovimento
    PreencheCboCartao
    lCxPeriodo = 0
    lCxCodigoLancamentoPadrao = 0
    lCxTipoMovimento = 2
    lCxCodigoUsuario = g_usuario
    lCxCodigoCartaoPreDefinido = 0
    cboCartao.Enabled = True
    txt_cartao.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub txt_cartao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboCartao.SetFocus
    End If
End Sub
Private Sub txt_cartao_LostFocus()
    Dim i As Integer
    If Val(txt_cartao.Text) > 0 And lOpcao > 0 Then
        If CartaoCredito.LocalizarCodigo(Val(txt_cartao.Text)) Then
            If TaxaAdmCartaoCredito.LocalizarCodigo(g_empresa, CartaoCredito.Codigo) Then
            Else
                TaxaAdmCartaoCredito.TaxaCusto = CartaoCredito.TaxaCusto
                MsgBox "Taxa de Adm de Cartão de crédito não cadastrada.", vbInformation, "Erro de Integridade!"
            End If
            'lbl_numero_lancamento.Caption = MovCartaoCredito.ProximoRegistro(g_empresa, CDate(txtDataEmissao.Text), cbo_periodo.Text)
            lbl_numero_lancamento.Caption = MovCartaoCredito.ProximoRegistro(g_empresa, CDate(txtDataEmissao.Text))
            cboCartao.ListIndex = -1
            Call SelecionaCartaoNaCombo(Val(txt_cartao.Text))
            
'            For i = 0 To cboCartao.ListCount - 1
'                If cboCartao.ItemData(i) = Val(txt_cartao.Text) Then
'                    cboCartao.ListIndex = i
'                    Exit For
'                End If
'            Next
            If txtValor.Enabled = True Then
                txtValor.SetFocus
            Else
                txtAutorizacao.SetFocus
            End If
        Else
            MsgBox "Cartão de crédito não cadastrado.", vbInformation, "Erro de Verificação!"
            txt_cartao.SetFocus
        End If
    End If
End Sub
Private Sub PreencheCboCartao()
    Dim rsCartao As ADODB.Recordset
    Dim xSQL As String
    
    Set rsCartao = New ADODB.Recordset
    xSQL = "SELECT Codigo, Nome FROM Cartao_Credito ORDER BY Nome"
    Set rsCartao = Conectar.RsConexao(xSQL)
    cboCartao.Clear
    With rsCartao
        If .RecordCount > 0 Then
            Do Until .EOF
                cboCartao.AddItem rsCartao("Nome").Value
                cboCartao.ItemData(cboCartao.NewIndex) = rsCartao("Codigo").Value
                .MoveNext
            Loop
        End If
    End With
    rsCartao.Close
    Set rsCartao = Nothing
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
Private Sub txt_numero_cartao_GotFocus()
    If lOpcao = 1 And txt_numero_cartao = "" Then
        txt_numero_cartao = 1
    End If
    txt_numero_cartao.SelStart = 0
    txt_numero_cartao.SelLength = Len(txt_numero_cartao)
End Sub
Private Sub txt_numero_cartao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtAutorizacao.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub zzLancaCartaoPelaComposicao()
    Dim rsComposicaoCaixa As New ADODB.Recordset
    Dim rsMovimentoComposicaoCaixa As New ADODB.Recordset
    Dim xSQL As String
    'Prepara SQL
    xSQL = ""
    xSQL = xSQL & "SELECT Codigo, Nome, Configuracao "
    xSQL = xSQL & "  FROM Composicao_Caixa"
    xSQL = xSQL & " WHERE MID(Configuracao,1,3) = " & Chr(39) & "CAR" & Chr(39)
    'Abre RecordSet
    Set rsComposicaoCaixa = New ADODB.Recordset
    Set rsComposicaoCaixa = Conectar.RsConexao(xSQL)
    'Verifica Composicao_Caixa
    If rsComposicaoCaixa.RecordCount > 0 Then
        Do Until rsComposicaoCaixa.EOF
        
        
        
            MsgBox rsComposicaoCaixa("Configuracao").Value
            'Prepara SQL
            xSQL = ""
            xSQL = xSQL & "SELECT *"
            xSQL = xSQL & "  FROM Movimento_Composicao_Caixa"
            xSQL = xSQL & " WHERE [Codigo da Composicao] = " & rsComposicaoCaixa("Codigo").Value
            xSQL = xSQL & "   AND Data >= #06/01/2004#"
            xSQL = xSQL & "   AND Data <= #12/31/2004#"
            'Abre RecordSet
            Set rsMovimentoComposicaoCaixa = New ADODB.Recordset
            Set rsMovimentoComposicaoCaixa = Conectar.RsConexao(xSQL)
        
            If rsMovimentoComposicaoCaixa.RecordCount > 0 Then
                Do Until rsMovimentoComposicaoCaixa.EOF
                    'VERIFICAR TOTAL
                    If MovCartaoCredito.TotalPeriodoBaixado(g_empresa, rsMovimentoComposicaoCaixa("Data").Value, rsMovimentoComposicaoCaixa("Periodo").Value, rsMovimentoComposicaoCaixa("Tipo do Movimento").Value, Val(Mid(rsComposicaoCaixa("Configuracao").Value, 4, 2))) = 0 Then
                        'Grava
                        If CartaoCredito.LocalizarCodigo(Val(Mid(rsComposicaoCaixa("Configuracao").Value, 4, 2))) Then
                            MovCartaoCredito.Empresa = rsMovimentoComposicaoCaixa("Empresa").Value
                            MovCartaoCredito.DataEmissao = rsMovimentoComposicaoCaixa("Data").Value
                            MovCartaoCredito.Periodo = rsMovimentoComposicaoCaixa("Periodo").Value
                            MovCartaoCredito.TipoMovimento = rsMovimentoComposicaoCaixa("Tipo do Movimento").Value
                            MovCartaoCredito.NumeroLancamento = 1
                            MovCartaoCredito.CodigoCartao = CartaoCredito.Codigo
                            MovCartaoCredito.DataVencimento = rsMovimentoComposicaoCaixa("Data").Value + CartaoCredito.DiasPrazo
                            MovCartaoCredito.Valor = rsMovimentoComposicaoCaixa("Valor").Value
                            MovCartaoCredito.NumeroCartao = "111"
                            MovCartaoCredito.Nome = ""
                            MovCartaoCredito.NumeroMovimentoCaixa = 0
                            MovCartaoCredito.TaxaAdministrativa = CartaoCredito.TaxaCusto
                            
                            If Not MovCartaoCredito.Incluir Then
                                MsgBox "Erro ao incluir Movimento de Cartao"
                            End If
                        Else
                            MsgBox "Cartao Não Cadstrado - " & rsComposicaoCaixa("Configuracao").Value
                        End If
                    End If
                rsMovimentoComposicaoCaixa.MoveNext
                Loop
            End If
            rsMovimentoComposicaoCaixa.Close
            rsComposicaoCaixa.MoveNext
        Loop
    End If
    rsComposicaoCaixa.Close
End Sub
Private Sub ProcessaCartaoCredito()
    Dim xData As Date
    On Error GoTo FileError
    
    xData = CDate("01/10/2004")
    If MovCartaoCredito.LocalizarPrimeiro Then
        If MovCartaoCredito.DataEmissao >= xData Then
            AtualizaTela
            If Not IncluiMovimentoCaixa Then
                MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
            Else
                MovCartaoCredito.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
                If Not MovCartaoCredito.Alterar(g_empresa, lData, lPeriodo, lOrdem) Then
                    MsgBox "Erro ao alterar cartão de Crédito", vbInformation, "Erro"
                End If
            End If
        End If
    
        Do Until MovCartaoCredito.LocalizarProximo = False
            If MovCartaoCredito.DataEmissao >= xData Then
                AtualizaTela
                If Not IncluiMovimentoCaixa Then
                    MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
                Else
                    MovCartaoCredito.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
                    If Not MovCartaoCredito.Alterar(g_empresa, lData, lPeriodo, lOrdem) Then
                        MsgBox "Erro ao alterar cartão de Crédito", vbInformation, "Erro"
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
    MsgBox "Erro ao processar Cartão de Crédito", vbInformation, "ProcessaCartaoCredito"
End Sub
Private Sub txtAutorizacao_GotFocus()
    txtAutorizacao.SelStart = 0
    txtAutorizacao.SelLength = Len(txtAutorizacao.Text)
End Sub
Private Sub txtAutorizacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtNSU.Enabled = True Then
            txtNSU.SetFocus
        Else
            cmd_ok.SetFocus
        End If
    End If
End Sub
Private Sub txtDataEmissao_GotFocus()
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
        If txtValor.Enabled = True Then
            txtValor.SetFocus
        Else
            txtAutorizacao.SetFocus
        End If
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtDataVencimento_LostFocus()
    txtDataVencimento.MaxLength = 10
    txtDataVencimento.Text = fMascaraData(txtDataVencimento.Text)
End Sub
Private Sub txtNSU_GotFocus()
    txtNSU.SelStart = 0
    txtNSU.SelLength = Len(txtNSU.Text)
End Sub
Private Sub txtNSU_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txtValor_GotFocus()
    If lOpcao = 1 Then
        cbo_tipo_movimento_LostFocus
    End If
    txtValor.SelStart = 0
    txtValor.SelLength = Len(txtValor.Text)
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_numero_cartao.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txtValor_LostFocus()
    If Val(txtValor.Text) > 0 Then
        txtValor.Text = Format(txtValor.Text, "###,##0.00")
    End If
End Sub
