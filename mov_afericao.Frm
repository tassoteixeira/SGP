VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form movimento_afericao 
   Caption         =   "Movimentação de Aferição/Devolução/Transferência"
   ClientHeight    =   6495
   ClientLeft      =   1785
   ClientTop       =   1500
   ClientWidth     =   6990
   Icon            =   "mov_afericao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "mov_afericao.frx":030A
   ScaleHeight     =   6495
   ScaleWidth      =   6990
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "mov_afericao.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Cria um novo registro."
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "mov_afericao.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Altera o registro atual."
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "mov_afericao.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Exclui o registro atual."
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "mov_afericao.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "mov_afericao.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5520
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   5355
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton btnSelecionaAbastecimento 
         Caption         =   "Abasteci&mentos..."
         Height          =   255
         Left            =   4200
         TabIndex        =   33
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton btnAtualizaGrid 
         Caption         =   "Listar de Abastecimentos Selecionados"
         Height          =   495
         Left            =   2160
         TabIndex        =   32
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox txtQuantidade 
         Height          =   285
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   14
         Top             =   1620
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cboIlha 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmd_transfere_dados_lmc 
         Caption         =   "&Transfere p/ LMC"
         Height          =   735
         Left            =   3240
         Picture         =   "mov_afericao.frx":7472
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Transfere as entradas de combustíveis para o LMC."
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txt_valor_total 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   16
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox chk_transferencia 
         Caption         =   "&Transferência"
         Height          =   255
         Left            =   4920
         TabIndex        =   17
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txt_valor_litro 
         Height          =   285
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbo_bomba 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
         Width           =   2475
      End
      Begin MSGrid.Grid gridAbastecimentoAfericao 
         Height          =   2055
         Left            =   120
         TabIndex        =   31
         Top             =   3120
         Width           =   6495
         _Version        =   65536
         _ExtentX        =   11456
         _ExtentY        =   3625
         _StockProps     =   77
         BackColor       =   16777215
         Cols            =   11
         FixedCols       =   0
      End
      Begin VB.Label Label3 
         Caption         =   "I&lha"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lbl_total 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5520
         TabIndex        =   12
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Total"
         Height          =   255
         Left            =   4920
         TabIndex        =   11
         Top             =   1620
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "&Valor do Litro"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Va&lor total"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "&Quantidade"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1620
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "&Período"
         Height          =   255
         Left            =   4920
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "&Data do movimento"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "&Bomba"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4680
      TabIndex        =   26
      Top             =   5400
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "mov_afericao.frx":8864
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "mov_afericao.frx":9D5E
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "mov_afericao.frx":B258
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "mov_afericao.frx":C6CA
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6060
      Picture         =   "mov_afericao.frx":DC4C
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Cancela o registro atual."
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5160
      Picture         =   "mov_afericao.frx":F146
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Confirma o registro atual."
      Top             =   5520
      Width           =   795
   End
End
Attribute VB_Name = "movimento_afericao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagMovimento As Integer
Dim lOpcao As Integer
Dim lData As Date
Dim lPeriodo As Integer
Dim lBomba As Integer
Dim lTipoCombustivel As String
Dim lTransferencia As Boolean
Dim lQuantidade As Currency
Dim lPrimeiraInclusao As Boolean
Dim lNumeroMovimentoCaixa As Long
Dim lIntegraCaixa As Boolean
Dim lIlha As Integer
Dim lValorAnterior As Currency
Dim lCaixaIndividual As Boolean

Dim lCxData As Date
Dim lCxPeriodo As String
Dim lCxTipoMovimento As Integer
Dim lCxIlha As Integer
Dim lCxBomba As Integer
Dim lCxDataDigitacao As Date
Dim lCxHoraDigitacao As Date
Dim lCxCodigoLancamentoPadrao As Integer
Dim lCxCodigoUsuario As Integer
Dim lCxCodigoFuncionario As Integer


'---------- Abastecimento Vinculado -----------
Dim lPossuiAbastecimentosVinculados As Boolean
'----------------------------------------------

Const NIVEL_ACESSO_DESENVOLVIMENTO As Integer = 1
Dim lNivelAcessoRestricaoLancarAfericaoSemAbstecimento As Integer

Private rsAuxiliar As New adodb.Recordset
Private Bomba As New cBomba
Private Combustivel As New cCombustivel
Private Configuracao As New cConfiguracao
Private ConfiguracaoDiversa As New cConfiguracaoDiversa
Private IntegracaoCaixa As New cIntegracaoCaixa
Private LivroLMC As New cLivroLMC
Private MovCaixaPista As New cMovimentoCaixaPista
Private MovAfericao As New cMovimentoAfericao
Private Abastecimento As New cMovimentoAbastecimento

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
    lCxCodigoLancamentoPadrao = Val(RetiraString(7, xString))

    lData = lCxData
    lPeriodo = lCxPeriodo
    lIlha = lCxIlha

    txtData.Enabled = False
    cbo_periodo.Enabled = False
    cboIlha.Enabled = False
    
    If xOperacao = "Incluir" Then
        lCxCodigoUsuario = Val(RetiraString(8, xString))
        lCxCodigoFuncionario = Val(RetiraString(9, xString))
        cmd_novo_Click
    ElseIf xOperacao = "Alterar" Then
        lCxBomba = CLng(RetiraString(8, xString))
        lCxCodigoUsuario = Val(RetiraString(9, xString))
        lCxCodigoFuncionario = Val(RetiraString(10, xString))
        lBomba = lCxBomba
        If MovAfericao.LocalizarCodigo(g_empresa, lData, lPeriodo, lBomba, lTransferencia) Then
            AtualTela
            If g_automacao And lPossuiAbastecimentosVinculados Then
                If (MsgBox("A aferição possui abastecimentos vinculados e não pode ser alterada" & vbCrLf & "Deseja excluir para lançar novamente?", 4 + 32 + 256, "Aferição com abastecimentos vinculados!")) = vbYes Then
                    cmd_excluir_Click
                    Exit Sub
                End If
            End If
        End If
        cmd_alterar_Click
    ElseIf xOperacao = "Excluir" Then
        lCxBomba = CLng(RetiraString(8, xString))
        lCxCodigoUsuario = Val(RetiraString(9, xString))
        lCxCodigoFuncionario = Val(RetiraString(10, xString))
        lBomba = lCxBomba
        If MovAfericao.LocalizarCodigo(g_empresa, lData, lPeriodo, lBomba, lTransferencia) Then
            AtualTela
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
Private Sub HabilitaDesabilitaControtes()


    If g_automacao = True And lOpcao = 1 Then
        btnSelecionaAbastecimento.Visible = True
        btnSelecionaAbastecimento.Enabled = True
        txt_valor_litro.Enabled = False
        txtQuantidade.Enabled = False
        txt_valor_total.Enabled = False
    Else
        btnSelecionaAbastecimento.Visible = False
        btnSelecionaAbastecimento.Enabled = False
        txt_valor_litro.Enabled = True
        txtQuantidade.Enabled = True
        txt_valor_total.Enabled = True
    End If

    If g_nivel_acesso = NIVEL_ACESSO_DESENVOLVIMENTO Then
        txt_valor_litro.Enabled = True
        txtQuantidade.Enabled = True
        txt_valor_total.Enabled = True
        cbo_bomba.Enabled = True
    ElseIf g_nivel_acesso <= lNivelAcessoRestricaoLancarAfericaoSemAbstecimento Then
        txt_valor_litro.Enabled = True
        txtQuantidade.Enabled = True
        txt_valor_total.Enabled = True
        cbo_bomba.Enabled = True
    End If

End Sub

Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_pesquisa.Enabled = False
    cmd_excluir.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    btnSelecionaAbastecimento.Enabled = False
End Sub
Private Sub AtualTabe()
    MovAfericao.Empresa = g_empresa
    MovAfericao.Data = CDate(txtData.Text)
    MovAfericao.Periodo = Val(cbo_periodo.Text)
    MovAfericao.CodigoBomba = cbo_bomba.ItemData(cbo_bomba.ListIndex)
    MovAfericao.Quantidade = fValidaValor2(txtQuantidade.Text)
    MovAfericao.PrecoCusto = Bomba.PrecoCusto
    MovAfericao.PrecoVenda = fValidaValor4(txt_valor_litro.Text)
    MovAfericao.ValorTotal = fValidaValor2(txt_valor_total.Text)
    MovAfericao.TipoCombustivel = Bomba.TipoCombustivel
    MovAfericao.Transferencia = chk_transferencia.Value
    MovAfericao.NumeroMovimentoCaixaCredito = lNumeroMovimentoCaixa
    MovAfericao.NumeroMovimentoCaixaDebito = lNumeroMovimentoCaixa
    MovAfericao.NumeroIlha = Val(cboIlha.Text)
    
 'ATENÇÃO - IMPLEMENTAÇÃO DE VINCULAR ABASTECIMENTO DEVERÁ SER REFEITA PARA MULTIPLOS ABASTECIMENTOS
    
    MovAfericao.DataAbastecimento = CDate("00:00:00")
    MovAfericao.HoraAbastecimento = CDate("00:00:00")
    
 
'    If g_automacao Then
'        MovAfericao.DataAbastecimento = CDate(txtDataAbastecimento.Text)
'        MovAfericao.HoraAbastecimento = CDate(txtHoraAbastecimento.Text)
'    Else
'        MovAfericao.DataAbastecimento = CDate("00:00:00")
'        MovAfericao.HoraAbastecimento = CDate("00:00:00")
'    End If
    
End Sub
Private Sub AtualTela()
    Dim i As Integer
    
    lData = MovAfericao.Data
    lPeriodo = MovAfericao.Periodo
    lIlha = MovAfericao.NumeroIlha
    lBomba = MovAfericao.CodigoBomba
    lTipoCombustivel = MovAfericao.TipoCombustivel
    lQuantidade = MovAfericao.Quantidade
    lTransferencia = MovAfericao.Transferencia
    lNumeroMovimentoCaixa = MovAfericao.NumeroMovimentoCaixaDebito
    lValorAnterior = MovAfericao.ValorTotal
    
    txtData.Text = Format(MovAfericao.Data, "dd/mm/yyyy")
    cbo_periodo.ListIndex = MovAfericao.Periodo - 1
    cboIlha.ListIndex = MovAfericao.NumeroIlha - 1
    
    Call SelecioneBombaNaCombo(MovAfericao.Empresa, MovAfericao.CodigoBomba)
    
    txt_valor_litro.Text = Format(MovAfericao.PrecoVenda, "###,##0.0000")
    txtQuantidade.Text = Format(MovAfericao.Quantidade, "###,##0.00")
    txt_valor_total.Text = Format(MovAfericao.ValorTotal, "###,##0.00")
    If MovAfericao.Transferencia Then
        chk_transferencia.Value = 1
    Else
        chk_transferencia.Value = 0
    End If
    
    If g_automacao Then
        Call AtualizaGridAbastecimentoAfericaoGravado(g_empresa, lData, lBomba, lPeriodo, lCxCodigoFuncionario)
        If lPossuiAbastecimentosVinculados Then
            cmd_alterar.Enabled = False
        End If
    End If
    
    
    lbl_total.Caption = Format(MovAfericao.TotalPeriodo(g_empresa, CDate(txtData.Text), Val(cbo_periodo.Text), MovAfericao.Transferencia), "###,##0.00")
    frm_dados.Enabled = False
    VerificaLiberacaoDigitacao
    Call VerificaLiberacaoLMC(lTipoCombustivel, lData)
End Sub

Private Sub SelecioneBombaNaCombo(ByVal pEmpresa As Integer, ByVal pCodigoBomba As Integer)
    Dim i As Integer
    If Bomba.LocalizarCodigo(pEmpresa, pCodigoBomba) Then
        For i = 0 To cbo_bomba.ListCount - 1
            cbo_bomba.ListIndex = i
            If cbo_bomba.ItemData(i) = Bomba.Codigo Then
                Exit For
            End If
        Next
    Else
        cbo_bomba.ListIndex = -1
    End If

End Sub

Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    FinalizaProcessoCaixa
    Set Bomba = Nothing
    Set Combustivel = Nothing
    Set Configuracao = Nothing
    Set ConfiguracaoDiversa = Nothing
    Set IntegracaoCaixa = Nothing
    Set LivroLMC = Nothing
    Set MovCaixaPista = Nothing
    Set MovAfericao = Nothing
    Set Abastecimento = Nothing
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    
    Call HabilitaDesabilitaControtes
End Sub
Function IncluiMovimentoCaixa() As Boolean
    Dim xComplemento As String
    IncluiMovimentoCaixa = False
    lNumeroMovimentoCaixa = 0
    If lIntegraCaixa = False Then
        IncluiMovimentoCaixa = True
        Exit Function
    End If
    If IntegracaoCaixa.LocalizarNome(g_empresa, "AFERIÇÃO") Then
        xComplemento = txtQuantidade.Text & " lts bc:" & Mid(cbo_bomba.Text, 1, 23) & " per:" & Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
        
        MovCaixaPista.Empresa = g_empresa
        MovCaixaPista.Data = CDate(txtData.Text)
        MovCaixaPista.NumeroMovimento = 1
        MovCaixaPista.Valor = fValidaValor(txt_valor_total.Text)
        MovCaixaPista.NumeroDocumento = Val(cbo_bomba.Text)
        MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
        MovCaixaPista.Complemento = xComplemento
        MovCaixaPista.NumeroContaDebito = IntegracaoCaixa.ContaDebito
        MovCaixaPista.NumeroContaCredito = IntegracaoCaixa.ContaCredito
        MovCaixaPista.TipoMovimento = lCxTipoMovimento
        MovCaixaPista.CodigoUsuario = lCxCodigoUsuario
        MovCaixaPista.Periodo = Val(cbo_periodo.Text)
        MovCaixaPista.NumeroIlha = Val(cboIlha.Text)
        MovCaixaPista.DadosInterno = "AFERI|@|" & Val(cbo_bomba.Text) & "|@|"
        MovCaixaPista.CodigoLancamentoPadrao = lCxCodigoLancamentoPadrao
        If lOpcao = 1 Then
            MovCaixaPista.DataDigitacao = Format(Date, "dd/MM/yyyy")
            MovCaixaPista.HoraDigitacao = Format(Time, "HH:mm:ss")
            MovCaixaPista.DataAlteracao = "00:00:00"
            MovCaixaPista.HoraAlteracao = "00:00:00"
        Else
            MovCaixaPista.DataDigitacao = lCxDataDigitacao
            MovCaixaPista.HoraDigitacao = lCxHoraDigitacao
            MovCaixaPista.DataAlteracao = Format(Date, "dd/MM/yyyy")
            MovCaixaPista.HoraAlteracao = Format(Time, "HH:mm:ss")
        End If
        If MovCaixaPista.Incluir Then
            IncluiMovimentoCaixa = True
            lNumeroMovimentoCaixa = MovCaixaPista.NumeroMovimento
        Else
            MsgBox "Não foi integrado no caixa o valor=" & txt_valor_total.Text, vbInformation, "Erro de Integridade"
        End If
    Else
        MsgBox "Não existe a integração=" & "AFERIÇÃO" & ".", vbInformation, "Registro Inexistente"
    End If
End Function
Private Sub IncluiSaidaBomba(pTipoCombustivel As String, pQuantidade As Currency)
    If Combustivel.LocalizarCodigo(g_empresa, pTipoCombustivel) Then
        Combustivel.QuantidadeEmEstoque = Combustivel.QuantidadeEmEstoque - pQuantidade
        If Not Combustivel.Alterar(g_empresa, pTipoCombustivel) Then
            MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Integridade"
        End If
    End If
End Sub
Private Sub ExcluiMovimentoCaixa()
    If lIntegraCaixa = False Then
        Exit Sub
    End If
    If MovCaixaPista.LocalizarCodigo(g_empresa, lData, lNumeroMovimentoCaixa) Then
        lCxDataDigitacao = MovCaixaPista.DataDigitacao
        lCxHoraDigitacao = MovCaixaPista.HoraDigitacao
        If Not MovCaixaPista.Excluir(g_empresa, lData, lNumeroMovimentoCaixa) Then
            MsgBox "Não foi excluído o movimento do caixa!", vbInformation, "Erro de Integridade."
        End If
    Else
        MsgBox "Não foi possível localizar o movimento do caixa!", vbInformation, "Erro de Integridade."
    End If
End Sub
Private Sub ExcluiSaidaBomba(pTipoCombustivel As String, pQuantidade As Currency)
    If Combustivel.LocalizarCodigo(g_empresa, pTipoCombustivel) Then
        Combustivel.QuantidadeEmEstoque = Combustivel.QuantidadeEmEstoque + pQuantidade
        If Not Combustivel.Alterar(g_empresa, pTipoCombustivel) Then
            MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Integridade"
        End If
    End If
End Sub
Private Sub PreencheCboBomba()
    Dim nome_combustivel As String
    Dim rsBomba As adodb.Recordset
    Dim xSQL As String
    
    Set rsBomba = New adodb.Recordset
    xSQL = "SELECT Codigo, [Tipo de Combustivel] FROM Bomba WHERE Empresa = " & g_empresa & " ORDER BY Codigo"
    Set rsBomba = Conectar.RsConexao(xSQL)
    cbo_bomba.Clear
    With rsBomba
        If .RecordCount > 0 Then
            Do Until .EOF
                If Combustivel.LocalizarCodigo(g_empresa, rsBomba("Tipo de Combustivel").Value) Then
                    nome_combustivel = Combustivel.Nome
                Else
                    nome_combustivel = "Não Cadastrado"
                End If
                cbo_bomba.AddItem rsBomba("Codigo").Value & " - " & nome_combustivel
                cbo_bomba.ItemData(cbo_bomba.NewIndex) = rsBomba("Codigo").Value
                .MoveNext
            Loop
        End If
    End With
    rsBomba.Close
    Set rsBomba = Nothing
End Sub



Private Sub AdcionaDadosGrid(lRS As adodb.Recordset)
    Dim x_i As Integer


    gridAbastecimentoAfericao.Row = gridAbastecimentoAfericao.Rows - 1
    gridAbastecimentoAfericao.Col = 0
    gridAbastecimentoAfericao.Text = Format(CDate(lRS("Data").Value), "dd/MM/yyyy")
    gridAbastecimentoAfericao.Col = 1
    gridAbastecimentoAfericao.Text = Format(CDate(lRS("Hora").Value), "HH:mm:ss")
    gridAbastecimentoAfericao.Col = 2
    gridAbastecimentoAfericao.Text = lRS("Tipo de Combustivel").Value
    gridAbastecimentoAfericao.Col = 3
    gridAbastecimentoAfericao.Text = Format(lRS("Valor Unitario").Value, "###,##0.0000")
    gridAbastecimentoAfericao.Col = 4
    gridAbastecimentoAfericao.Text = Format(lRS("Quantidade").Value, "###,##0.00")
    gridAbastecimentoAfericao.Col = 5
    gridAbastecimentoAfericao.Text = Format(lRS("Valor Total").Value, "###,##0.00")
    gridAbastecimentoAfericao.Col = 6
    gridAbastecimentoAfericao.Text = lRS("Codigo do Funcionario").Value
    gridAbastecimentoAfericao.Col = 7
    gridAbastecimentoAfericao.Text = lRS("Codigo do Produto").Value
    gridAbastecimentoAfericao.Rows = gridAbastecimentoAfericao.Rows + 1
End Sub


Private Sub MontaGrid()

    LimpaGrid
    gridAbastecimentoAfericao.Row = 0
    gridAbastecimentoAfericao.Col = 0
    gridAbastecimentoAfericao.Text = "Dt.Abast."
    gridAbastecimentoAfericao.ColWidth(0) = TextWidth(String$(10, "9"))
    gridAbastecimentoAfericao.ColAlignment(0) = 2
   'obs: o "9"equivale ao tab
    '0 = left, 1 = right ,2 =  center
    gridAbastecimentoAfericao.Col = 1
    gridAbastecimentoAfericao.Text = "Hr.Abast."
    gridAbastecimentoAfericao.ColWidth(1) = TextWidth(String$(10, "9"))
    gridAbastecimentoAfericao.ColAlignment(1) = 2
    gridAbastecimentoAfericao.Col = 2
    gridAbastecimentoAfericao.Text = "Comb."
    gridAbastecimentoAfericao.ColWidth(2) = TextWidth(String$(4, "9"))
    gridAbastecimentoAfericao.ColAlignment(2) = 2
    gridAbastecimentoAfericao.Col = 3
    gridAbastecimentoAfericao.Text = "Vl.Unit."
    gridAbastecimentoAfericao.ColWidth(3) = TextWidth(String$(10, "9"))
    gridAbastecimentoAfericao.ColAlignment(3) = 1
    gridAbastecimentoAfericao.Col = 4
    gridAbastecimentoAfericao.Text = "Qtd."
    gridAbastecimentoAfericao.ColWidth(4) = TextWidth(String$(10, "9"))
    gridAbastecimentoAfericao.ColAlignment(4) = 1
    gridAbastecimentoAfericao.Col = 5
    gridAbastecimentoAfericao.Text = "Vl.Total"
    gridAbastecimentoAfericao.ColWidth(5) = TextWidth(String$(10, "9"))
    gridAbastecimentoAfericao.ColAlignment(5) = 1
    gridAbastecimentoAfericao.Col = 6
    gridAbastecimentoAfericao.Text = "Func."
    gridAbastecimentoAfericao.ColWidth(6) = TextWidth(String$(10, "9"))
    gridAbastecimentoAfericao.ColAlignment(6) = 0
    gridAbastecimentoAfericao.Col = 7
    gridAbastecimentoAfericao.Text = "Prod."
    gridAbastecimentoAfericao.ColWidth(7) = TextWidth(String$(8, "9"))
    gridAbastecimentoAfericao.ColAlignment(7) = 1
'    gridAbastecimentoAfericao.Col = 8
'    gridAbastecimentoAfericao.Text = "Ordem Dig."
'    gridAbastecimentoAfericao.ColWidth(8) = TextWidth(String$(6, "9"))
'    gridAbastecimentoAfericao.ColAlignment(8) = 1
'    gridAbastecimentoAfericao.Col = 9
'    gridAbastecimentoAfericao.Text = "Codigo Prod"
'    gridAbastecimentoAfericao.ColWidth(9) = TextWidth(String$(6, "9"))
'    gridAbastecimentoAfericao.ColAlignment(9) = 1
'    gridAbastecimentoAfericao.Col = 10
'    gridAbastecimentoAfericao.Text = "Tipo Mov."
'    gridAbastecimentoAfericao.ColWidth(9) = TextWidth(String$(4, "9"))
'    gridAbastecimentoAfericao.ColAlignment(9) = 1
End Sub
Sub LimpaGrid()
    Do Until gridAbastecimentoAfericao.Rows = 2
        gridAbastecimentoAfericao.Row = gridAbastecimentoAfericao.Rows - 1
        gridAbastecimentoAfericao.RemoveItem gridAbastecimentoAfericao.Row
    Loop
    gridAbastecimentoAfericao.Row = 1
    gridAbastecimentoAfericao.Col = 0
    gridAbastecimentoAfericao.Text = ""
    gridAbastecimentoAfericao.Col = 1
    gridAbastecimentoAfericao.Text = ""
    gridAbastecimentoAfericao.Col = 2
    gridAbastecimentoAfericao.Text = ""
    gridAbastecimentoAfericao.Col = 3
    gridAbastecimentoAfericao.Text = ""
    gridAbastecimentoAfericao.Col = 4
    gridAbastecimentoAfericao.Text = ""
    gridAbastecimentoAfericao.Col = 5
    gridAbastecimentoAfericao.Text = ""
    gridAbastecimentoAfericao.Col = 6
    gridAbastecimentoAfericao.Text = ""
    gridAbastecimentoAfericao.Col = 7
    gridAbastecimentoAfericao.Text = ""
'    gridAbastecimentoAfericao.Col = 8
'    gridAbastecimentoAfericao.Text = ""
'    gridAbastecimentoAfericao.Col = 9
'    gridAbastecimentoAfericao.Text = ""
'    gridAbastecimentoAfericao.Col = 10
'    gridAbastecimentoAfericao.Text = ""
End Sub


Private Sub btnAtualizaGrid_Click()
    Call GravaAuditoria(1, Me.name, 23, btnSelecionaAbastecimento.Caption & " Func.:" & lCxCodigoFuncionario)
    If lOpcao = 1 Then
        Call AtualizaGridAbastecimentoAfericaoPendente(g_empresa, CDate(txtData.Text), Val(cbo_bomba.Text), Val(cbo_periodo.Text), lCxCodigoFuncionario)
    Else
        Call AtualizaGridAbastecimentoAfericaoGravado(g_empresa, lData, lBomba, lPeriodo, lCxCodigoFuncionario)
    End If
End Sub
Private Sub LiberarAbastecimentoAfericaoPendente(ByVal pData As Date, ByVal pCodigoFuncionario As Integer)
    Dim xMovimentoAbastecimento As New CadastroDLL.cMovimentoAbastecimento
    
    Call xMovimentoAbastecimento.LiberaAbastecimentoAfericaoTemporaria(g_empresa, pData, pCodigoFuncionario)
    
End Sub

Public Function AtualizaGridAbastecimentoAfericaoPendente(ByVal pEmpresa As Integer, pData As Date, ByVal pBico As Integer, ByVal pPeriodo As Integer, ByVal pCodigoFuncionario As Integer) As adodb.Recordset

On Error GoTo trata_erro
    
    Dim xRsMovAbastecimento As New adodb.Recordset
    Dim xSQL As String
    Dim xTotal As Currency
    Dim xQuanidadeTotal As Currency
    Dim xValorUnitario As Currency

    MontaGrid

    
    xSQL = "SELECT * FROM Movimento_Abastecimento"
    xSQL = xSQL & " WHERE Empresa = " & pEmpresa
    xSQL = xSQL & " AND Bico = " & pBico
    xSQL = xSQL & " AND Data = " & preparaData(pData)
    xSQL = xSQL & " AND Periodo = " & pPeriodo
    xSQL = xSQL & " AND [Complemento Documento Gerado] = " & preparaTexto("AFTEMP")
    If pCodigoFuncionario > 0 Then
        xSQL = xSQL & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    End If
    
    xRsMovAbastecimento.CursorLocation = adUseClient
    Set xRsMovAbastecimento = Conectar.RsConexao(xSQL)

    If xRsMovAbastecimento.RecordCount > 0 Then
       lPossuiAbastecimentosVinculados = True
       Do Until xRsMovAbastecimento.EOF
          Call AdcionaDadosGrid(xRsMovAbastecimento)
          xTotal = xTotal + xRsMovAbastecimento("Valor Total").Value
          xQuanidadeTotal = xQuanidadeTotal + xRsMovAbastecimento("Quantidade").Value
          xRsMovAbastecimento.MoveNext
       Loop
    Else
        lPossuiAbastecimentosVinculados = False
    End If
    
    txt_valor_total.Text = Format(xTotal, "###,##0.00")
    txtQuantidade.Text = Format(xQuanidadeTotal, "###,##0.00")
    If xTotal > 0 And xQuanidadeTotal > 0 Then
        txt_valor_litro.Text = Format(xTotal / xQuanidadeTotal, "###,##0.0000")
    Else
        txt_valor_litro.Text = Format(0, "###,##0.0000")
    End If
    gridAbastecimentoAfericao.Row = gridAbastecimentoAfericao.Rows - 1
    gridAbastecimentoAfericao.Col = 0

    
    xRsMovAbastecimento.Close
    Set xRsMovAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description

End Function
Public Function AtualizaGridAbastecimentoAfericaoGravado(ByVal pEmpresa As Integer, pDataAfericao As Date, ByVal pBico As Integer, ByVal pPeriodo As Integer, ByVal pCodigoFuncionario As Integer) As adodb.Recordset

On Error GoTo trata_erro
    
    Dim xRsMovAbastecimento As New adodb.Recordset
    Dim xSQL As String

    MontaGrid

    
    xSQL = "SELECT [Empresa],[Bico],[Data da Afericao],[Periodo],[Data do Abastecimento] as Data,[Hora do Abastecimento] as Hora,[Codigo do Produto],[Tipo de Combustivel],[Valor Unitario],[Quantidade],[Valor Total],[Codigo do Funcionario]"
    xSQL = xSQL & " FROM AbastecimentoAfericaoAutomacao"
    xSQL = xSQL & " WHERE Empresa = " & pEmpresa
    xSQL = xSQL & " AND Bico = " & pBico
    xSQL = xSQL & " AND [Data da Afericao] = " & preparaData(pDataAfericao)
    xSQL = xSQL & " AND Periodo = " & pPeriodo
    
    If pCodigoFuncionario > 0 Then
        xSQL = xSQL & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    End If
    
    xRsMovAbastecimento.CursorLocation = adUseClient
    Set xRsMovAbastecimento = Conectar.RsConexao(xSQL)

    If xRsMovAbastecimento.RecordCount > 0 Then
       lPossuiAbastecimentosVinculados = True
       Do Until xRsMovAbastecimento.EOF
          Call AdcionaDadosGrid(xRsMovAbastecimento)
          xRsMovAbastecimento.MoveNext
       Loop
    Else
        lPossuiAbastecimentosVinculados = False
    End If
    
    gridAbastecimentoAfericao.Row = gridAbastecimentoAfericao.Rows - 1
    gridAbastecimentoAfericao.Col = 0

    
    xRsMovAbastecimento.Close
    Set xRsMovAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description

End Function



Private Sub btnSelecionaAbastecimento_Click()
    
    If cbo_bomba.ListIndex < 0 Then
       MsgBox "A bomba deve ser selecionada", vbInformation, "Campo obrigtório."
       cbo_bomba.SetFocus
       Exit Sub
    ElseIf cbo_periodo.ListIndex < 0 Then
       MsgBox "O período deve ser informado", vbInformation, "Campo obrigtório."
       cbo_periodo.SetFocus
       Exit Sub
    End If


    Call GravaAuditoria(1, Me.name, 23, btnSelecionaAbastecimento.Caption & " Func.:" & lCxCodigoFuncionario)
 

    If lCxPeriodo > 0 Then
        Call ChamaConsultaAbastecimento
    Else
        gStringChamada = Val(cbo_bomba.ItemData(cbo_bomba.ListIndex)) & "|@|" & lCxCodigoFuncionario & "|@|" & Val(cbo_periodo.Text) & "|@|"
        Call menu_personalizado.GravaSgpNetCadastroIni("ConsultaAbastecimentoAfericaoAutomacao")
    End If


    Call AguardeSelecaoAbastecimento '- IMPLEMENTAR

End Sub
Public Sub ChamaConsultaAbastecimento()
    Dim xArquivoDiscoTMP As TextStream
    
    On Error GoTo FileError
    
    g_string = "SELECIONAR_ABASTECIMENTO_AFERICAO|@|" & Val(cbo_bomba.ItemData(cbo_bomba.ListIndex)) & "|@|" & lCxCodigoFuncionario & "|@|" & Val(cbo_periodo.Text) & "|@|"
    
    Set xArquivoDiscoTMP = gArqTxt.CreateTextFile("C:" & gDiretorioAplicativo & "Retorno_VB6_Fim.TMP")
    xArquivoDiscoTMP.WriteLine ("[Outras]")
    xArquivoDiscoTMP.WriteLine ("gString=" & g_string)
    xArquivoDiscoTMP.Close
    
    Call gArqTxt.MoveFile("C:" & gDiretorioAplicativo & "Retorno_VB6_Fim.TMP", "C:" & gDiretorioAplicativo & "Retorno_VB6_Fim.INI")
    
    g_string = ""
    
    Exit Sub
FileError:
End Sub


Private Sub AguardeSelecaoAbastecimento()
'Implementar posteriormente


End Sub
Private Sub cbo_bomba_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtQuantidade.Enabled Then
            txtQuantidade.SetFocus
        End If
    End If
End Sub
Private Sub cbo_bomba_LostFocus()
    If cbo_bomba.ListIndex <> -1 Then
        If Bomba.LocalizarCodigo(g_empresa, Val(cbo_bomba.ItemData(cbo_bomba.ListIndex))) Then
            txt_valor_litro.Text = Format(Bomba.PrecoVenda, "###0.0000")
        Else
            txt_valor_litro.Text = 0
        End If
        If lOpcao = 1 Then
            If MovAfericao.LocalizarCodigo(g_empresa, CDate(txtData.Text), Val(cbo_periodo.Text), Val(cbo_bomba.Text), chk_transferencia.Value) Then
                MsgBox "Já existe aferição cadastrada com este código." & Chr(10) & Chr(10) & "Mude o código informado.", vbInformation, "Duplicidade de registro!"
                cbo_bomba.ListIndex = -1
                cbo_bomba.SetFocus
            End If
            If lPossuiAbastecimentosVinculados Then
                LimpaGrid
            End If
        End If
    End If
End Sub
Private Sub cbo_periodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboIlha.SetFocus
    End If
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
Private Sub cboIlha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If g_automacao Then
            btnSelecionaAbastecimento.SetFocus
        Else
            cbo_bomba.SetFocus
        End If
        cbo_bomba.SetFocus

    End If
End Sub
Private Sub chk_transferencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
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
    
    Call HabilitaDesabilitaControtes

    If g_automacao Then
        If btnSelecionaAbastecimento.Enabled Then
            btnSelecionaAbastecimento.SetFocus
        End If
    Else
        If txtQuantidade.Enabled Then
          txtQuantidade.SetFocus
        End If
    End If
    txtQuantidade.SetFocus

End Sub
Private Sub cmd_anterior_Click()
    Call GravaAuditoria(1, Me.name, 13, "")
    If MovAfericao.LocalizarAnterior Then
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
    If MovAfericao.LocalizarCodigo(g_empresa, lData, lPeriodo, lBomba, lTransferencia) Then
        AtualTela
        AtivaBotoes
        cmd_alterar.SetFocus
        If lPossuiAbastecimentosVinculados Then
            cmd_alterar.Enabled = False
        End If
    Else
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
        cmd_novo.SetFocus
    End If
    lOpcao = 0
End Sub
Private Sub LimpaTela()
    If lPrimeiraInclusao = False Then
        txtData.Text = ""
        cbo_periodo.ListIndex = -1
        cboIlha.ListIndex = -1
    End If
    cbo_bomba.ListIndex = -1
    txt_valor_litro.Text = ""
    txtQuantidade.Text = ""
    txt_valor_total.Text = ""
    chk_transferencia.Value = 0
    LimpaGrid
End Sub
Private Sub cmd_excluir_Click()
    Call GravaAuditoria(1, Me.name, 4, "")
    If cbo_bomba.ListIndex > -1 Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = vbYes Then
            Call GravaAuditoria(1, Me.name, 10, "Dt:" & txtData.Text & " Per:" & Val(cbo_periodo.Text) & " Ilha:" & Val(cboIlha.Text) & " Vlr:" & txt_valor_total.Text & " Bico:" & Val(cbo_bomba.Text))
            Call ExcluiMovimentoCaixa
            
            If MovAfericao.Excluir(g_empresa, lData, lPeriodo, lBomba, lTransferencia) Then
                If MovAfericao.Transferencia = True Then
                    Call ExcluiSaidaBomba(lTipoCombustivel, lQuantidade)
                End If
                If g_automacao And lPossuiAbastecimentosVinculados Then
                    Call DesvinculaAbastecimento(MovAfericao.Data, MovAfericao.CodigoBomba, MovAfericao.TipoCombustivel)
                End If
                LimpaTela
                If MovAfericao.LocalizarUltimo(g_empresa) Then
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
Private Sub VinculaAbastecimento()
    Dim xIndiceLinha As Integer
    
    If g_automacao Then
        If gridAbastecimentoAfericao.Rows > 2 Then
            For xIndiceLinha = 2 To gridAbastecimentoAfericao.Rows
                Dim xAbastecimentoAfericao As CadastroDLL.cAbastecimentoAfericaoAuto
                
                Set xAbastecimentoAfericao = PreencheAbastecimentoAfericaoLinhaGrid(xIndiceLinha)
                
                If Not xAbastecimentoAfericao Is Nothing Then
                    If Not xAbastecimentoAfericao.Incluir Then
                        MsgBox "Não foi possível vincular abastecimento - Data: " & xAbastecimentoAfericao.DataAbastecimento & " Hora: " & xAbastecimentoAfericao.HoraAbastecimento & " Comb.: " & xAbastecimentoAfericao.TipoCombustivel, vbOKOnly, "Erro de integridade"
                    Else
                        Call DefineAbastecimentoVinculado(xAbastecimentoAfericao.DataAbastecimento, xAbastecimentoAfericao.HoraAbastecimento, xAbastecimentoAfericao.Bico, xAbastecimentoAfericao.Periodo, xAbastecimentoAfericao.CodigoFuncionario)
                    End If
                End If
            Next
        End If
    Else
        Exit Sub
    End If
End Sub
Private Function PreencheAbastecimentoAfericaoLinhaGrid(ByVal pLinhaGrid As Integer) As CadastroDLL.cAbastecimentoAfericaoAuto
        Dim xIndiceColuna As Integer
        
        Dim xAbastecimentoAfericao As New CadastroDLL.cAbastecimentoAfericaoAuto
                        
        Set PreencheAbastecimentoAfericaoLinhaGrid = Nothing
                        
        xAbastecimentoAfericao.Empresa = g_empresa
        xAbastecimentoAfericao.DataAfericao = CDate(txtData.Text)
        xAbastecimentoAfericao.Bico = Val(cbo_bomba.Text)
        xAbastecimentoAfericao.Periodo = Val(cbo_periodo.Text)
        gridAbastecimentoAfericao.Row = pLinhaGrid - 1
        
        For xIndiceColuna = 0 To gridAbastecimentoAfericao.Cols - 1
            gridAbastecimentoAfericao.Col = xIndiceColuna
            
            If xIndiceColuna = 0 Then
                If Val(gridAbastecimentoAfericao.Text) = 0 Then
                    Exit Function
                End If
                xAbastecimentoAfericao.DataAbastecimento = CDate(gridAbastecimentoAfericao.Text)
            ElseIf xIndiceColuna = 1 Then
                xAbastecimentoAfericao.HoraAbastecimento = CDate(gridAbastecimentoAfericao.Text)
            ElseIf xIndiceColuna = 2 Then
                xAbastecimentoAfericao.TipoCombustivel = gridAbastecimentoAfericao.Text
            ElseIf xIndiceColuna = 3 Then
                xAbastecimentoAfericao.ValorUnitario = CCur(gridAbastecimentoAfericao.Text)
            ElseIf xIndiceColuna = 4 Then
                xAbastecimentoAfericao.Quantidade = CCur(gridAbastecimentoAfericao.Text)
            ElseIf xIndiceColuna = 5 Then
                xAbastecimentoAfericao.ValorTotal = CCur(gridAbastecimentoAfericao.Text)
            ElseIf xIndiceColuna = 6 Then
                xAbastecimentoAfericao.CodigoFuncionario = Val(gridAbastecimentoAfericao.Text)
            ElseIf xIndiceColuna = 7 Then
                xAbastecimentoAfericao.CodigoProduto = Val(gridAbastecimentoAfericao.Text)
            End If
        Next

        Set PreencheAbastecimentoAfericaoLinhaGrid = xAbastecimentoAfericao
End Function


Private Sub DefineAbastecimentoVinculado(ByVal pDataAbastecimento As Date, ByVal pHoraAbastecimento As Date, ByVal pBico As Integer, ByVal pPeriodo As Integer, ByVal pCodigoFuncionario As Integer)
    
    If Not Abastecimento.DefineAbastecimentoVinculadoAfericao(g_empresa, pDataAbastecimento, pHoraAbastecimento, pBico, pPeriodo, pCodigoFuncionario) Then
        MsgBox "Não foi possível definir abastecimento para aferição", vbOKOnly, "Erro de integridade"
    End If
    
End Sub
Private Sub DesvinculaAbastecimento(ByVal pData As Date, ByVal pBico As Integer, ByVal pTipoCombustivel As String)
    
    Dim xIndiceLinha As Integer
    
    If g_automacao Then
        If gridAbastecimentoAfericao.Rows > 2 Then
            For xIndiceLinha = 2 To gridAbastecimentoAfericao.Rows
                Dim xAbastecimentoAfericao As CadastroDLL.cAbastecimentoAfericaoAuto
                
                Set xAbastecimentoAfericao = PreencheAbastecimentoAfericaoLinhaGrid(xIndiceLinha)
                If Not xAbastecimentoAfericao Is Nothing Then
                    If Not xAbastecimentoAfericao.Excluir(xAbastecimentoAfericao.Empresa, xAbastecimentoAfericao.Bico, xAbastecimentoAfericao.DataAfericao, xAbastecimentoAfericao.DataAbastecimento, xAbastecimentoAfericao.HoraAbastecimento, xAbastecimentoAfericao.CodigoFuncionario) Then
                        MsgBox "Não foi possível desvincular abastecimento - Data: " & xAbastecimentoAfericao.DataAbastecimento & " Hora: " & xAbastecimentoAfericao.HoraAbastecimento & " Comb.: " & xAbastecimentoAfericao.TipoCombustivel, vbOKOnly, "Erro de integridade"
                    Else
                        Call Abastecimento.LiberaAbastecimentoAfericao(xAbastecimentoAfericao.Empresa, xAbastecimentoAfericao.DataAbastecimento, xAbastecimentoAfericao.HoraAbastecimento, xAbastecimentoAfericao.Bico, xAbastecimentoAfericao.CodigoFuncionario)
                    End If
                End If
            Next
        End If
    Else
        Exit Sub
    End If
End Sub

Private Sub cmd_novo_Click()
    Call GravaAuditoria(1, Me.name, 2, "")
    LimpaTela
    frm_dados.Enabled = True
    Inclui
    If lCxPeriodo > 0 Then
        txtData.Text = Format(lCxData, "dd/mm/yyyy")
        cbo_periodo.ListIndex = lCxPeriodo - 1
        cboIlha.ListIndex = lCxIlha - 1
        
        cbo_bomba.SetFocus
'        If cbo_bomba.Enabled Then
'           cbo_bomba.SetFocus
'        Else
'           btnSelecionaAbastecimento.SetFocus
'        End If
        
        Exit Sub
    Else
        If lPrimeiraInclusao = False Then
            txtData.SetFocus
        Else
            cbo_bomba.SetFocus
'            If cbo_bomba.Enabled Then
'               cbo_bomba.SetFocus
'            Else
'btnSelecionaAbastecimento.SetFocus
'            End If
        End If
    End If
End Sub
Private Sub cmd_novo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 16 Then
        MsgBox "PROCESSAMENTO"
        Call ProcessaAfericao
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        If VerificaLiberacaoDigitacao2 Then
            AtivaBotoes
            If lOpcao = 1 Then
                If Not IncluiMovimentoCaixa Then
                    MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
                End If
                AtualTabe
                Call GravaAuditoria(1, Me.name, 10, "Dt:" & txtData.Text & " Per:" & Val(cbo_periodo.Text) & " Ilha:" & Val(cboIlha.Text) & " Vlr:" & txt_valor_total.Text & " Bico:" & Val(cbo_bomba.Text))
                If MovAfericao.Incluir Then
                    lData = MovAfericao.Data
                    lPeriodo = MovAfericao.Periodo
                    lIlha = MovAfericao.NumeroIlha
                    lBomba = MovAfericao.CodigoBomba
                    lTransferencia = MovAfericao.Transferencia
                    lPrimeiraInclusao = True
                    Call VinculaAbastecimento
                    'lRegistro = Val(txt_registro.Text)
                Else
                    MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
                End If
                If chk_transferencia.Value = 1 Then
                    Call IncluiSaidaBomba(lTipoCombustivel, lQuantidade)
                End If
            ElseIf lOpcao = 2 Then
            
                If g_automacao And lPossuiAbastecimentosVinculados Then
                    MsgBox "Operação de alteração não permitida, o registro aferição deve ser excluído e cadastrado novamente.", vbInformation, "Operação não permitida."
                    Exit Sub
                End If
            
                Call ExcluiMovimentoCaixa
                
                If MovAfericao.Transferencia = True Then
                    Call ExcluiSaidaBomba(lTipoCombustivel, lQuantidade)
                End If
                If Not IncluiMovimentoCaixa Then
                    MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
                End If
                
                AtualTabe
                Call GravaAuditoria(1, Me.name, 10, "Para: Dt:" & lData & " Per:" & lPeriodo & " Ilha:" & lIlha & " Vlr:" & lValorAnterior & " Bico:" & lBomba)
                Call GravaAuditoria(1, Me.name, 10, "Para: Dt:" & txtData.Text & " Per:" & Val(cbo_periodo.Text) & " Ilha:" & Val(cboIlha.Text) & " Vlr:" & txt_valor_total.Text & " Bico:" & Val(cbo_bomba.Text))
                If MovAfericao.Alterar(g_empresa, lData, lPeriodo, lBomba, lTransferencia) Then
                    lData = MovAfericao.Data
                    lPeriodo = MovAfericao.Periodo
                    lIlha = MovAfericao.NumeroIlha
                    lBomba = MovAfericao.CodigoBomba
                    'Call VinculaAbastecimento(MovAfericao.DataAbastecimento, MovAfericao.HoraAbastecimento, MovAfericao.CodigoBomba, MovAfericao.TipoCombustivel)
                Else
                    MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
                End If
                If chk_transferencia.Value = 1 Then
                    Call IncluiSaidaBomba(lTipoCombustivel, lQuantidade)
                End If
                If lCxPeriodo > 0 Then
                    cmd_sair_Click
                    Exit Sub
                End If
            End If
            If MovAfericao.LocalizarCodigo(g_empresa, lData, lPeriodo, lBomba, lTransferencia) Then
                AtualTela
            Else
                LimpaTela
            End If
            If lOpcao = 1 Then
                cmd_novo_Click
            Else
                lOpcao = 0
                cmd_alterar.SetFocus
            End If
        End If
    End If
    Exit Sub
FileError:
    MsgBox Error
    Exit Sub
End Sub
Private Sub TransfereDadosLMC()
    Dim x_data As Date
    Dim xSQL As String
    
    On Error GoTo FileError
    
    x_data = CDate("01/01/1900")
    'Busca ultima data com movimento
    If MovAfericao.LocalizarUltimo(g_empresa) Then
        x_data = MovAfericao.Data
    End If
    x_data = x_data + 1
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será transferido a aferição de combustível apartir da data " & x_data & "." & Chr(10) & Chr(10) & "Deseja realmente fazer esta transferência?", vbYesNo + 256, "Transfere a Aferição de Combustível Para o L.M.C.!")) = vbNo Then
        Exit Sub
    End If
    Call GravaAuditoria(1, Me.name, 10, "Empresa:" & g_empresa & " A Partir de:" & x_data)
    
    'Seleciona Registros a Serem Transferidos
    xSQL = ""
    xSQL = xSQL & "   SELECT Data, Periodo, [Codigo da Bomba], Quantidade, [Preco de Custo],"
    xSQL = xSQL & "          [Preco de Venda], [Valor Total], [Tipo de Combustivel], Transferencia,"
    xSQL = xSQL & "          [Numero do Movimento do Caixa Credito], [Numero do Movimento do Caixa Debito],"
    xSQL = xSQL & "          [Numero da Ilha]"
    'xSQL = xSQL & "          [Numero da Ilha], [Data do Abastecimento], [Hora do Abastecimento]"
    xSQL = xSQL & "     FROM Movimento_Afericao"
    xSQL = xSQL & "    WHERE Empresa = " & g_empresa
    xSQL = xSQL & "      AND Data >= " & preparaData(x_data)
    xSQL = xSQL & " ORDER BY Data, Periodo"
    Set rsAuxiliar = New adodb.Recordset
    Set rsAuxiliar = Conectar.RsConexao(xSQL)
    
    'Transfere Dados para o LMC
    If Not rsAuxiliar.EOF Then
        Do Until rsAuxiliar.EOF
            MovAfericao.Empresa = g_empresa
            MovAfericao.Data = rsAuxiliar("Data").Value
            MovAfericao.Periodo = rsAuxiliar("Periodo").Value
            MovAfericao.CodigoBomba = rsAuxiliar("Codigo da Bomba").Value
            MovAfericao.Quantidade = rsAuxiliar("Quantidade").Value
            MovAfericao.PrecoCusto = rsAuxiliar("Preco de Custo").Value
            MovAfericao.PrecoVenda = rsAuxiliar("Preco de Venda").Value
            MovAfericao.ValorTotal = rsAuxiliar("Valor Total").Value
            MovAfericao.TipoCombustivel = rsAuxiliar("Tipo de Combustivel").Value
            MovAfericao.Transferencia = rsAuxiliar("Transferencia").Value
            MovAfericao.NumeroMovimentoCaixaCredito = rsAuxiliar("Numero do Movimento do Caixa Credito").Value
            MovAfericao.NumeroMovimentoCaixaDebito = rsAuxiliar("Numero do Movimento do Caixa Debito").Value
            MovAfericao.NumeroIlha = rsAuxiliar("Numero da Ilha").Value
            
            'vão ser removidos do objeto MovAfericao
            MovAfericao.DataAbastecimento = CDate("00:00:00")
            MovAfericao.HoraAbastecimento = CDate("00:00:00")
            
            If Not MovAfericao.Incluir Then
                MsgBox "Registro não foi gravado!", vbInformation, "Erro de Integridade"
            End If
            rsAuxiliar.MoveNext
        Loop
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com a aferição de combustível transferida para o L.M.C.", vbInformation, "Transferência Concluida!"
    Else
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Não tem aferição de combustível à ser transferida para o L.M.C.", vbInformation, "Transferência Não Concluida!"
    End If
    rsAuxiliar.Close
    Set rsAuxiliar = Nothing
    'If MovAfericao.LocalizarUltimo(g_empresa) Then
    '    lData = MovAfericao.Data
    'End If
    Exit Sub

FileError:
    MsgBox Error
    Resume Next

End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(txtData.Text) Then
        MsgBox "Informe a data.", vbInformation, "Atenção!"
        txtData.SetFocus
    ElseIf cbo_periodo.ListIndex = -1 Then
        MsgBox "Selecione o periodo.", vbInformation, "Atenção!"
        cbo_periodo.SetFocus
    ElseIf cboIlha.ListIndex = -1 Then
        MsgBox "Selecione uma ilha.", vbInformation, "Atenção!"
        cboIlha.SetFocus
    ElseIf cbo_bomba.ListIndex = -1 Then
        MsgBox "Selecione a bomba.", vbInformation, "Atenção!"
        cbo_bomba.SetFocus
'        If cbo_bomba.Enabled Then
'           cbo_bomba.SetFocus
'        Else
           'btnSelecionaAbastecimento.SetFocus
'        End If

    ElseIf Not fValidaValor4(txt_valor_litro.Text) > 0 Then
        MsgBox "Informe o valor do litro.", vbInformation, "Atenção!"
        
        txt_valor_litro.SetFocus
'        If txt_valor_litro.Enabled Then
'           txt_valor_litro.SetFocus
'        Else
'           'btnSelecionaAbastecimento.SetFocus
'        End If

    ElseIf Not fValidaValor2(txtQuantidade.Text) > 0 Then
        MsgBox "Informe a quantidade.", vbInformation, "Atenção!"
        txtQuantidade.SetFocus
'        If txtQuantidade.Enabled Then
'           txtQuantidade.SetFocus
'        Else
'           'btnSelecionaAbastecimento.SetFocus
'        End If
    ElseIf Not fValidaValor2(txt_valor_total.Text) > 0 Then
        MsgBox "Informe o valor total.", vbInformation, "Atenção!"
        txt_valor_total.SetFocus
'        If txt_valor_total.Enabled Then
'           txt_valor_total.SetFocus
'        Else
'           'btnSelecionaAbastecimento.SetFocus
'        End If
        
    ElseIf Not VerificaCriarAfericao() Then
           cbo_bomba.SetFocus
'        If cbo_bomba.Enabled Then
'           cbo_bomba.SetFocus
'        Else
'           'btnSelecionaAbastecimento.SetFocus
'        End If
    Else
        ValidaCampos = True
    End If
End Function
Private Sub VerificaLiberacaoDigitacao()
    Dim x_flag As Boolean
    x_flag = True
    If g_nivel_acesso > 4 Then
        If MovAfericao.Empresa < g_cfg_empresa_i Or MovAfericao.Empresa > g_cfg_empresa_f Then
            x_flag = False
        ElseIf MovAfericao.Data < g_cfg_data_i Or MovAfericao.Data > g_cfg_data_f Then
            x_flag = False
        ElseIf MovAfericao.Periodo < g_cfg_periodo_i Or MovAfericao.Periodo > g_cfg_periodo_f Then
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
    
    If g_automacao And lPossuiAbastecimentosVinculados Then
       cmd_alterar.Enabled = False
    End If

End Sub
Function VerificaLiberacaoDigitacao2() As Boolean
    VerificaLiberacaoDigitacao2 = False
    If g_nivel_acesso <= 4 Then
        VerificaLiberacaoDigitacao2 = True
        Exit Function
    End If
    If IsDate(lCxData) Then
        If lCxData = CDate(txtData.Text) Then
            VerificaLiberacaoDigitacao2 = True
            Exit Function
        End If
    End If
    If CDate(txtData.Text) < g_cfg_data_i Or CDate(txtData.Text) > g_cfg_data_f Then
        MsgBox "A data do movimento deve estar entre " & Format(g_cfg_data_i, "dd/mm/yyyy") & " a " & Format(g_cfg_data_f, "dd/mm/yyyy") & ".", vbInformation, "Digitação Não Autorizada!"
        txtData.SetFocus
    ElseIf cbo_periodo.Text < g_cfg_periodo_i Or cbo_periodo.Text > g_cfg_periodo_f Then
        MsgBox "O período deve estar entre " & g_cfg_periodo_i & " ao " & g_cfg_periodo_f & ".", vbInformation, "Digitação Não Autorizada!"
        cbo_periodo.SetFocus
    Else
        VerificaLiberacaoDigitacao2 = True
    End If
End Function
Function VerificaCriarAfericao() As Boolean
    VerificaCriarAfericao = True
    
    If ConfiguracaoDiversa.LocalizarCodigo(1, "RESTRINGE:CRIAR AFERICAO") Then
        If ConfiguracaoDiversa.Codigo <= g_nivel_acesso Then
            MsgBox "Este usuário não está autorizado a realizar esta operação", vbInformation, "Digitação Não Autorizada!"
            VerificaCriarAfericao = False
        End If
    End If
End Function

Private Function VerificaLiberacaoLMC(ByVal pTipoCombustivel As String, ByVal pData As Date) As Boolean
    If g_nome_usuario = "L.M.C." Then
        VerificaLiberacaoLMC = False
        If LivroLMC.LocalizarCombustivelConcluido(g_empresa, pTipoCombustivel, pData) = "NAO" Then
            VerificaLiberacaoLMC = True
            cmd_alterar.Enabled = True
            cmd_excluir.Enabled = True
        Else
            cmd_alterar.Enabled = False
            cmd_excluir.Enabled = False
        End If
    Else
        VerificaLiberacaoLMC = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    Call GravaAuditoria(1, Me.name, 5, "")
    consulta_movimento_afericao.Show 1
    If Len(g_string) > 0 Then
        lData = RetiraGString(1)
        lPeriodo = RetiraGString(2)
        lBomba = RetiraGString(3)
        If RetiraGString(4) = "Yes" Then
            lTransferencia = True
        Else
            lTransferencia = False
        End If
        If MovAfericao.LocalizarCodigo(g_empresa, lData, lPeriodo, lBomba, lTransferencia) Then
            AtualTela
        Else
            LimpaTela
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    Call GravaAuditoria(1, Me.name, 12, "")
    If MovAfericao.LocalizarPrimeiro() Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    Call GravaAuditoria(1, Me.name, 14, "")
    If MovAfericao.LocalizarProximo Then
        AtualTela
    Else
        MsgBox "Fim de Arquivo.", vbInformation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_transfere_dados_lmc_Click()
    Call GravaAuditoria(1, Me.name, 23, "Transferencia Para LMC")
    If MovAfericao.TransfereDadosLMC(g_empresa, True) Then
        Call GravaAuditoria(1, Me.name, 10, "Empresa:" & g_empresa & " A Partir de:" & MovAfericao.UltimaData(g_empresa))
        If MovAfericao.TransfereDadosLMC(g_empresa, False) Then
            MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com a aferição de combustível transferida para o L.M.C.", vbInformation, "Transferência Concluida!"
        Else
            MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Não tem aferição de combustível à ser transferida para o L.M.C.", vbInformation, "Transferência Não Concluida!"
        End If
    Else
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Não tem aferição de combustível à ser transferida para o L.M.C.", vbInformation, "Transferência Não Concluida!"
    End If
    'TransfereDadosLMC
    cmd_cancelar_Click
    AtivaBotoes
    cmd_ultimo_Click
End Sub
Private Sub cmd_ultimo_Click()
    Call GravaAuditoria(1, Me.name, 15, "")
    If MovAfericao.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If lFlagMovimento = 0 Then
        DesativaBotoes
        If RetiraGString(1) = "CaixaPista" Then
            AjustaCaixaPista
        Else
            If MovAfericao.LocalizarUltimo(g_empresa) Then
                AtivaBotoes
                AtualTela
                If g_automacao And lPossuiAbastecimentosVinculados Then
                    cmd_alterar.Enabled = False
                End If
                Call VerificaLiberacaoLMC(lTipoCombustivel, lData)
            Else
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
            End If
            If cmd_novo.Enabled Then
                cmd_novo.SetFocus
            End If
        End If
    Else
        lFlagMovimento = 0
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
    Screen.MousePointer = 1
    CentraForm Me
    lCaixaIndividual = False
    If ConfiguracaoDiversa.LocalizarCodigo(g_empresa, "CAIXA DE PISTA INDIVIDUAL") Then
        lCaixaIndividual = ConfiguracaoDiversa.Verdadeiro
    End If
    If g_nome_usuario = "L.M.C." Then
        MovAfericao.NomeTabela = "Movimento_Afericao_LMC"
        Me.Caption = Me.Caption & " - LMC"
        lIntegraCaixa = False
    Else
        MovAfericao.NomeTabela = "Movimento_Afericao"
        lIntegraCaixa = True
    End If
    
    
    If ConfiguracaoDiversa.LocalizarCodigo(1, "RESTRINGE:LANÇAR AFERICAO SEM ABAST.") Then
        lNivelAcessoRestricaoLancarAfericaoSemAbstecimento = ConfiguracaoDiversa.Codigo
    Else
        lNivelAcessoRestricaoLancarAfericaoSemAbstecimento = NIVEL_ACESSO_DESENVOLVIMENTO
    End If
   
    
    PreencheCboBomba
    PreencheCboIlha
    PreencheCboPeriodo
    lPrimeiraInclusao = False
    lCxPeriodo = 0
    lCxCodigoLancamentoPadrao = 4
    lCxTipoMovimento = 2
    lCxCodigoUsuario = g_usuario
    lCxCodigoFuncionario = 0
    lPossuiAbastecimentosVinculados = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub txtData_GotFocus()
    If Not IsDate(txtData.Text) And lOpcao = 1 Then
        txtData.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
    End If
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
Private Sub txt_valor_litro_GotFocus()
    txt_valor_litro.SelStart = 0
    txt_valor_litro.SelLength = Len(txt_valor_litro.Text)
End Sub
Private Sub txt_valor_litro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtQuantidade.SetFocus
    End If
End Sub
Private Sub txt_valor_litro_LostFocus()
    txt_valor_litro.Text = Format(txt_valor_litro.Text, "###,##0.0000")
End Sub
Private Sub txt_valor_total_GotFocus()
    txt_valor_total.SelStart = 0
    txt_valor_total.SelLength = Len(txt_valor_total.Text)
End Sub
Private Sub txt_valor_total_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_valor_total_LostFocus()
    txt_valor_total.Text = Format(txt_valor_total.Text, "###,##0.00")
End Sub
Private Sub ProcessaAfericao()
    Dim xData As Date
    On Error GoTo FileError
    
    xData = CDate("01/10/2004")
    If MovAfericao.LocalizarPrimeiro Then
        If MovAfericao.Data >= xData Then
            AtualTela
            If Not IncluiMovimentoCaixa Then
                MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
            Else
                MovAfericao.NumeroMovimentoCaixaCredito = lNumeroMovimentoCaixa
                MovAfericao.NumeroMovimentoCaixaDebito = lNumeroMovimentoCaixa
                If Not MovAfericao.Alterar(g_empresa, lData, lPeriodo, lBomba, lTransferencia) Then
                    MsgBox "Erro ao alterar aferição", vbInformation, "Erro"
                End If
            End If
        End If
    
        Do Until MovAfericao.LocalizarProximo = False
            If MovAfericao.Data >= xData Then
                AtualTela
                If Not IncluiMovimentoCaixa Then
                    MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
                Else
                    MovAfericao.NumeroMovimentoCaixaCredito = lNumeroMovimentoCaixa
                    MovAfericao.NumeroMovimentoCaixaDebito = lNumeroMovimentoCaixa
                    If Not MovAfericao.Alterar(g_empresa, lData, lPeriodo, lBomba, lTransferencia) Then
                        MsgBox "Erro ao alterar aferição", vbInformation, "Erro"
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
    MsgBox "Erro ao processar aferição", vbInformation, "ProcessaAfericao"
End Sub
Private Sub txtQuantidade_GotFocus()
    txtQuantidade.SelStart = 0
    txtQuantidade.SelLength = Len(txtQuantidade.Text)
End Sub
Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_total.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txtQuantidade_LostFocus()
    If Trim(txtQuantidade.Text) <> "" Then
        txtQuantidade.Text = Format(txtQuantidade.Text, "###,##0.00")
        txt_valor_total.Text = Format(fValidaValor2(txtQuantidade.Text) * fValidaValor4(txt_valor_litro.Text), "###,##0.00")
    End If
End Sub
