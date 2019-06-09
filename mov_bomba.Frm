VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form movimento_bomba 
   Caption         =   "Movimentação das Bombas"
   ClientHeight    =   5415
   ClientLeft      =   1785
   ClientTop       =   1500
   ClientWidth     =   6975
   Icon            =   "mov_bomba.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5415
   ScaleWidth      =   6975
   Begin VB.Timer TimerAutomacao 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   60
      Top             =   4920
   End
   Begin MSCommLib.MSComm MSCommAutomacao 
      Left            =   720
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "mov_bomba.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Cria um novo registro."
      Top             =   4440
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "mov_bomba.frx":199C
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Altera o registro atual."
      Top             =   4440
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "mov_bomba.frx":2E96
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Exclui o registro atual."
      Top             =   4440
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "mov_bomba.frx":4528
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   4440
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "mov_bomba.frx":599A
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4440
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.TextBox txtData 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtNumeroSubCaixa 
         Height          =   285
         Left            =   4980
         MaxLength       =   3
         TabIndex        =   6
         Top             =   720
         Width           =   435
      End
      Begin VB.CheckBox chkSubCaixa 
         Caption         =   "SubCai&xa N."
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmd_encerrante 
         Caption         =   "Ca&lcula Encerrante"
         Height          =   795
         Left            =   4860
         Picture         =   "mov_bomba.frx":702C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Calcula o encerrante de bomba."
         Top             =   2580
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CommandButton cmd_transfere_dados_lmc 
         Caption         =   "&Transfere p/ LMC"
         Height          =   795
         Left            =   4860
         Picture         =   "mov_bomba.frx":7336
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Transfere o movimento de bombas para o LMC."
         Top             =   1620
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox txtQuantidade 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   16
         Top             =   3060
         Width           =   1335
      End
      Begin VB.TextBox txtValorLitro 
         Height          =   285
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   10
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txt_tanque 
         Height          =   285
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   20
         Top             =   3840
         Width           =   255
      End
      Begin VB.TextBox txtEncerrante 
         Height          =   285
         Left            =   1680
         MaxLength       =   13
         TabIndex        =   14
         ToolTipText     =   "Tecle CTRL+R p/ repetir a abertura."
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtAbertura 
         Height          =   285
         Left            =   1680
         MaxLength       =   13
         TabIndex        =   12
         Top             =   2160
         Width           =   1335
      End
      Begin VB.ComboBox cbo_bomba 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1200
         Width           =   4935
      End
      Begin VB.Label Label7 
         Caption         =   "Tanque"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "&Valor do Litro"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblValorSaida 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Valor da Saida"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "&Quantidade da Saida"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3120
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "&Período"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
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
         Caption         =   "&Bomba"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "&Abertura"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "&Encerrante"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   1455
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4680
      TabIndex        =   41
      Top             =   4320
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "mov_bomba.frx":8728
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "mov_bomba.frx":9C22
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "mov_bomba.frx":B11C
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "mov_bomba.frx":C58E
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5160
      Picture         =   "mov_bomba.frx":DB10
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Confirma o registro atual."
      Top             =   4440
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6060
      Picture         =   "mov_bomba.frx":F11A
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Cancela o registro atual."
      Top             =   4440
      Width           =   795
   End
   Begin VB.Frame frm_calcula_encerrante 
      Caption         =   "Calcula Encerrante"
      Height          =   2355
      Left            =   1680
      TabIndex        =   23
      Top             =   2040
      Visible         =   0   'False
      Width           =   3915
      Begin VB.TextBox txtDataInicial 
         Height          =   285
         Left            =   180
         MaxLength       =   10
         TabIndex        =   25
         Top             =   540
         Width           =   1095
      End
      Begin VB.TextBox txtDataFinal 
         Height          =   285
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   27
         Top             =   540
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cancelar2 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2940
         TabIndex        =   33
         ToolTipText     =   "Cancela o cálculo."
         Top             =   1800
         Width           =   795
      End
      Begin VB.CommandButton cmd_ok2 
         Caption         =   "&OK"
         Height          =   375
         Left            =   1920
         TabIndex        =   32
         ToolTipText     =   "Confirma o cálculo desta bomba."
         Top             =   1800
         Width           =   795
      End
      Begin VB.TextBox txt_encerrante_inicial 
         Height          =   285
         Left            =   180
         MaxLength       =   12
         TabIndex        =   31
         Top             =   1860
         Width           =   1335
      End
      Begin VB.ComboBox cbo_bomba_2 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1200
         Width           =   3555
      End
      Begin VB.Label Label13 
         Caption         =   "Encerrante inicial"
         Height          =   255
         Left            =   180
         TabIndex        =   30
         Top             =   1620
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Bomba"
         Height          =   255
         Left            =   180
         TabIndex        =   28
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Data final"
         Height          =   255
         Left            =   2700
         TabIndex        =   26
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Data inicial"
         Height          =   255
         Left            =   180
         TabIndex        =   24
         Top             =   300
         Width           =   975
      End
   End
End
Attribute VB_Name = "movimento_bomba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_movimento_bomba As Integer
Dim lOpcao As Integer
Dim lData As Date
Dim lPeriodo As Integer
Dim lBomba As Integer
Dim lIlha As Integer
Dim lNumeroSubCaixa As Integer
Dim lTipoCombustivel As String
Dim lQuantidade As Currency
Dim lTentativaLeitura As Integer

Dim l_qtd_periodo As Integer
Dim lQuantidadeBico As Integer
Dim l_ultima_ilha As Integer
Dim l_pular_bomba As Integer
Dim lUltimaData As Date
Dim lEncerraMesmoPeriodo As Boolean
Dim lDataECF As Date
Dim lQtdA As Currency
Dim lQtdD As Currency
Dim lQtdG As Currency
Dim lTotalA As Currency
Dim lTotalD As Currency
Dim lTotalG As Currency
Dim lNomeArquivo As String
Dim BemaRetorno As Integer
Dim lSerieECF As String
Dim lCodigoEcf As Integer
Dim lQuantidadeMaximaBico As Long

Dim lNivelAcessoRestringeEditarEncerranteAbertura As Integer

Dim lAutomacaoPorta As Integer
Dim lAutomacaoVelocidade As String
Dim lAutomacaoDtr As Boolean
Dim lAutomacaoRts As Boolean
Dim lAutomacaoFlag As Integer
Dim lAutomacaoBicoEncerrante As Integer
Dim lAutomacaoAbertura As Currency
Dim lAutomacaoEncerrante As Currency
Dim lAutomacaoData As Date
Dim lAutomacaoPeriodo As Integer
Dim lAutomacaoIlha As Integer

Private Bomba As New cBomba
Private Combustivel As New cCombustivel
Private Configuracao As New cConfiguracao
Private ConfiguracaoDiversa As New cConfiguracaoDiversa
Private ECF As New cEcf
Private EncerranteAtual As New cEncerranteAtual
Private FechamentoCaixa As New cFechamentoCaixa
Private LivroLMC As New cLivroLMC
Private MovimentoAbastecimento As New cMovimentoAbastecimento
Private MovimentoBomba As New cMovimentoBomba
Private MovimentoCupomFiscal As New cMovimentoCupomFiscal
Private Sub AtualizaConstantes()
    l_qtd_periodo = 1
    lQuantidadeBico = 1
    l_ultima_ilha = 1
    l_pular_bomba = 1
    If Configuracao.LocalizarCodigo(g_empresa) Then
        l_qtd_periodo = Configuracao.QuantidadePeriodos
        lQuantidadeBico = Configuracao.QuantidadeBico
        l_ultima_ilha = Configuracao.QuantidadeIlha
        l_pular_bomba = Configuracao.PularBomba
    End If
    lSerieECF = ReadINI("CUPOM FISCAL", "Serie ECF", gArquivoIni)
    lCodigoEcf = 1
    If ECF.LocalizarNumeroSerie(g_empresa, lSerieECF) Then
        lCodigoEcf = ECF.Codigo
    End If
    lQuantidadeMaximaBico = 5000
    If ConfiguracaoDiversa.LocalizarCodigo(1, "MOV BOMBA: Quantidade Maxima Por Bico") Then
        lQuantidadeMaximaBico = ConfiguracaoDiversa.Codigo
    End If
    
    lNivelAcessoRestringeEditarEncerranteAbertura = 5 'Inicio permitindo LMC acima
    
    If ConfiguracaoDiversa.LocalizarCodigo(1, "RESTRINGE:EDITAR ENCERRANTE ABERTURA") Then
        lNivelAcessoRestringeEditarEncerranteAbertura = ConfiguracaoDiversa.Codigo
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
    cmd_transfere_dados_lmc.Visible = False
    cmd_encerrante.Visible = False
    frm_move.Visible = True
End Sub
Function ConverteBico(ByVal pBico As String) As String
    ConverteBico = pBico
    If UCase(g_nome_empresa) Like "*BOSQUE*" Then
        If pBico = "01" Then
            ConverteBico = "08"
        ElseIf pBico = "02" Then
            ConverteBico = "09"
        ElseIf pBico = "03" Then
            ConverteBico = "04"
        ElseIf pBico = "04" Then
            ConverteBico = "05"
        ElseIf pBico = "05" Then
            ConverteBico = "0C"
        ElseIf pBico = "06" Then
            ConverteBico = "0D"
        ElseIf pBico = "07" Then
            ConverteBico = "06"
        ElseIf pBico = "08" Then
            ConverteBico = "07"
        End If
    End If
End Function
Private Sub DeletaMovimentoData(ByVal pData As Date)
    Dim xSQL As String
    xSQL = "DELETE FROM Movimento_Bomba_Cupom "
    xSQL = xSQL & "WHERE Empresa = " & g_empresa
    xSQL = xSQL & "  AND Data = " & preparaData(pData)
    Conectar.ExecutaSql (xSQL)
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
    If g_nome_usuario = "L.M.C." Then
        cmd_transfere_dados_lmc.Visible = True
        If g_lmc = 2 Then
            cmd_encerrante.Visible = True
        End If
    ElseIf UCase(g_nome_usuario) <> "CUPOM FISCAL" Or fEcfInstalada = False Then
        cmd_transfere_dados_lmc.Visible = True
    End If
End Sub
Private Sub AtualTabe()
    MovimentoBomba.Empresa = g_empresa
    MovimentoBomba.Data = CDate(txtData.Text)
    MovimentoBomba.Periodo = cbo_periodo.Text
    MovimentoBomba.SubCaixa = Val(txtNumeroSubCaixa.Text)
    MovimentoBomba.CodigoBomba = Mid(cbo_bomba.ItemData(cbo_bomba.ListIndex), 2, 2)
    MovimentoBomba.Abertura = Format(fValidaValor(txtAbertura.Text), "###,##0.000")
    MovimentoBomba.Encerrante = Format(fValidaValor(txtEncerrante.Text), "###,##0.000")
    MovimentoBomba.QuantidadeSaida = Format(fValidaValor(txtQuantidade.Text), "###,##0.000")
    MovimentoBomba.PrecoCusto = Bomba.PrecoCusto
    MovimentoBomba.PrecoVenda = fValidaValor4(txtValorLitro.Text)
    MovimentoBomba.TipoCombustivel = Bomba.TipoCombustivel
    MovimentoBomba.NumeroTanque = txt_tanque.Text
    MovimentoBomba.NumeroIlha = Mid(cbo_bomba.ItemData(cbo_bomba.ListIndex), 1, 1)
    MovimentoBomba.TotalVendaBruto = MovimentoBomba.QuantidadeSaida * MovimentoBomba.PrecoVenda
    MovimentoBomba.TotalDesconto = 0
    MovimentoBomba.TotalAcrescimo = 0

End Sub
Private Sub AtualTela()
    Dim i As Integer
    
    lData = MovimentoBomba.Data
    lPeriodo = MovimentoBomba.Periodo
    lBomba = MovimentoBomba.CodigoBomba
    lIlha = MovimentoBomba.NumeroIlha
    lTipoCombustivel = MovimentoBomba.TipoCombustivel
    lQuantidade = MovimentoBomba.QuantidadeSaida
    lNumeroSubCaixa = MovimentoBomba.SubCaixa
    
    txtData.Text = Format(MovimentoBomba.Data, "dd/mm/yyyy")
    cbo_periodo.ListIndex = MovimentoBomba.Periodo - 1
    txtNumeroSubCaixa.Text = MovimentoBomba.SubCaixa
    If MovimentoBomba.SubCaixa = 999 Then
        chkSubCaixa.Value = 0
        txtNumeroSubCaixa.Enabled = False
    Else
        chkSubCaixa.Value = 1
        txtNumeroSubCaixa.Enabled = True
    End If
    If Bomba.LocalizarCodigo(MovimentoBomba.Empresa, MovimentoBomba.CodigoBomba) Then
        For i = 0 To cbo_bomba.ListCount - 1
            cbo_bomba.ListIndex = i
            If cbo_bomba.ItemData(i) = Bomba.NumeroIlha & Format(Bomba.Codigo, "00") Then
                Exit For
            End If
        Next
    Else
        cbo_bomba.ListIndex = -1
    End If
    txtValorLitro.Text = Format(MovimentoBomba.PrecoVenda, "###,##0.0000")
    txtAbertura.Text = Format(MovimentoBomba.Abertura, "###,##0.000")
    txtEncerrante.Text = Format(MovimentoBomba.Encerrante, "###,##0.000")
    txtQuantidade.Text = Format(MovimentoBomba.QuantidadeSaida, "###,##0.000")
    lblValorSaida.Caption = Format(fValidaValor(txtQuantidade.Text) * MovimentoBomba.PrecoVenda, "###,##0.00")
    txt_tanque.Text = MovimentoBomba.NumeroTanque
    
    frm_dados.Enabled = False
    VerificaLiberacaoDigitacao
    Call VerificaLiberacaoLMC(lTipoCombustivel, lData)
End Sub
Public Sub AutomacaoAbrePorta()
    If MSCommAutomacao.PortOpen = False Then
        MSCommAutomacao.CommPort = lAutomacaoPorta
        MSCommAutomacao.Settings = lAutomacaoVelocidade
        MSCommAutomacao.PortOpen = True
    End If
End Sub
Private Sub AutomacaoAtualTabe()
    Dim sql As String
    'sql = "Delete * From Movimento_Bomba "
    'sql = sql & "Where Empresa = " & g_empresa
    'sql = sql & " And Data = #02/24/2000#"
    'bd_sgp.Execute sql
    AutomacaoPreparaAberturaBomba
    If Bomba.LocalizarCodigo(g_empresa, lAutomacaoBicoEncerrante) Then
        If MovimentoBomba.LocalizarCodigo(g_empresa, lAutomacaoData, lAutomacaoPeriodo, lAutomacaoBicoEncerrante, 999) Then
            If lAutomacaoPeriodo < l_qtd_periodo Then
                If lAutomacaoBicoEncerrante = 1 Then
                    If (MsgBox("O período " & lAutomacaoPeriodo & " já foi encerrado." & Chr(13) & "Deseja encerrar o mesmo período novamente?", vbQuestion + vbYesNo + vbDefaultButton2, "Período já Encerrado!")) = vbYes Then
                        lEncerraMesmoPeriodo = True
                    Else
                        lEncerraMesmoPeriodo = False
                    End If
                End If
                If lEncerraMesmoPeriodo Then
                    If Not MovimentoBomba.Excluir(g_empresa, lAutomacaoData, lAutomacaoPeriodo, lAutomacaoBicoEncerrante, 999) Then
                        MsgBox "Erro ao excluir registro de mov. bomba!", vbInformation, "Erro de Integridade!"
                    End If
                    AutomacaoPreparaAberturaBomba
                Else
                    lAutomacaoPeriodo = lAutomacaoPeriodo + 1
                End If
            Else
                If Not MovimentoBomba.Excluir(g_empresa, lAutomacaoData, lAutomacaoPeriodo, lAutomacaoBicoEncerrante, 999) Then
                    MsgBox "Erro ao excluir registro de mov. bomba!", vbInformation, "Erro de Integridade!"
                End If
                AutomacaoPreparaAberturaBomba
            End If
        End If
        
        If MovimentoBomba.LocalizarCodigo(g_empresa, lAutomacaoData, lAutomacaoPeriodo, lAutomacaoBicoEncerrante, 999) Then
            If Not MovimentoBomba.Excluir(g_empresa, lAutomacaoData, lAutomacaoPeriodo, lAutomacaoBicoEncerrante, 999) Then
                MsgBox "Erro ao excluir registro de mov. bomba!", vbInformation, "Erro de Integridade!"
            End If
            AutomacaoPreparaAberturaBomba
            lAutomacaoPeriodo = lAutomacaoPeriodo + 1
        End If
        
        MovimentoBomba.Empresa = g_empresa
        MovimentoBomba.Data = lAutomacaoData
        MovimentoBomba.Periodo = lAutomacaoPeriodo
        MovimentoBomba.CodigoBomba = lAutomacaoBicoEncerrante
        If lAutomacaoAbertura = 0 Then
            lAutomacaoAbertura = lAutomacaoEncerrante
        End If
        MovimentoBomba.Abertura = lAutomacaoAbertura
        MovimentoBomba.Encerrante = lAutomacaoEncerrante
        MovimentoBomba.QuantidadeSaida = lAutomacaoEncerrante - lAutomacaoAbertura
        MovimentoBomba.PrecoCusto = Bomba.PrecoCusto
        MovimentoBomba.PrecoVenda = Bomba.PrecoVenda
        MovimentoBomba.TipoCombustivel = Bomba.TipoCombustivel
        MovimentoBomba.NumeroTanque = Bomba.NumeroTanque
        MovimentoBomba.NumeroIlha = Bomba.NumeroIlha
        MovimentoBomba.SubCaixa = 999
        MovimentoBomba.TotalVendaBruto = MovimentoBomba.QuantidadeSaida * MovimentoBomba.PrecoVenda
        MovimentoBomba.TotalDesconto = 0
        MovimentoBomba.TotalAcrescimo = 0

        
        lData = MovimentoBomba.Data
        lPeriodo = MovimentoBomba.Periodo
        lBomba = MovimentoBomba.CodigoBomba
        lIlha = MovimentoBomba.NumeroIlha
        lNumeroSubCaixa = MovimentoBomba.SubCaixa
        If Not MovimentoBomba.Incluir Then
            MsgBox "Não foi possível incluir o registro de movimento de bomba!", vbCritical, "Erro de Integridade!"
        End If
        If MovimentoBomba.LocalizarCodigo(g_empresa, lData, lPeriodo, lBomba, lNumeroSubCaixa) Then
            AtualTela
        End If
    Else
        MsgBox "Bomba não cadastrada!", vbInformation, "AutomacaoAtualTabe"
    End If
    lAutomacaoBicoEncerrante = lAutomacaoBicoEncerrante + 1
    If Not UCase(g_nome_empresa) Like "*J & R*" Then
        If lAutomacaoBicoEncerrante > lQuantidadeBico Then
            lAutomacaoBicoEncerrante = 0
            TimerAutomacao.Enabled = False
            GravaFechamentoCaixa
            'MsgBox "Leitura de encerrantes efetuada!", vbInformation, "Leitura OK!"
            cmd_sair_Click
        End If
    End If
End Sub
Public Sub AutomacaoBuscaEncerrantes(ByVal pBico As String)
    Dim i As Integer
    Dim x_string As String
    Dim x_quant As Integer
    
    'Le encerrante
    If UCase(g_nome_empresa) Like "*J & R*" Then
        lData = RetiraGString(4)
        lPeriodo = RetiraGString(5)
        If EncerranteAtual.LocalizarPrimeiro Then
            lAutomacaoAbertura = EncerranteAtual.Encerrante
            lAutomacaoEncerrante = EncerranteAtual.Encerrante
            AutomacaoAtualTabe
            Do Until EncerranteAtual.LocalizarProximo = True
                lAutomacaoAbertura = EncerranteAtual.Encerrante
                lAutomacaoEncerrante = EncerranteAtual.Encerrante
                AutomacaoAtualTabe
            Loop
        End If
        lAutomacaoBicoEncerrante = 0
        TimerAutomacao.Enabled = False
        GravaFechamentoCaixa
        'MsgBox "Leitura de encerrantes efetuada!", vbInformation, "Leitura OK!"
        cmd_sair_Click
    End If
    x_quant = MSCommAutomacao.InBufferCount
    If x_quant = 5 Then
        lTentativaLeitura = lTentativaLeitura + 1
        x_string = MSCommAutomacao.Input
        Call CriaLogCupom(Time & " - ERRO Movimento de Bomba: AutomacaoBuscaEncerrantes: x_string=" & x_string & " Tentativa=" & lTentativaLeitura)
        MsgBox "Não foi possível ler o encerrante.", vbInformation, "Bico " & Format(lAutomacaoBicoEncerrante, "00") & " em uso!"
    ElseIf x_quant = 16 Then
        lTentativaLeitura = 0
        x_string = MSCommAutomacao.Input
        Call CriaLogCupom(Time & " - Movimento de Bomba: AutomacaoBuscaEncerrantes: x_string=" & x_string)
        'MsgBox "Bico: " & Mid(x_string, 4, 2) & Chr(10) & "Encerrante: " & Mid(x_string, 6, 6) & "," & Mid(x_string, 12, 2), vbInformation, "Leitura de Encerrante!"
        lData = RetiraGString(4)
        lPeriodo = RetiraGString(5)
        lAutomacaoAbertura = Mid(x_string, 6, 6) & "," & Mid(x_string, 12, 2)
        lAutomacaoEncerrante = Mid(x_string, 6, 6) & "," & Mid(x_string, 12, 2)
        AutomacaoAtualTabe
    End If
End Sub
Public Sub AutomacaoLeStatusBomba()
    Dim x_quant As Integer
    Dim x_string As String
    
    
                lAutomacaoPorta = PortaDaAutomacao
                '----- Velocidade
                lAutomacaoVelocidade = "9600,n,8,1"
                '----- DTR
                lAutomacaoDtr = True
                'ldtr = False
                '----- RTS
                lAutomacaoRts = True
                'lRts = False
                '----- EOF
                AutomacaoAbrePorta
                lAutomacaoFlag = 0
    
    
    AutomacaoAbrePorta
    
    'Envia Comando para Pedir Status das bombas
    MSCommAutomacao.InBufferCount = 0
    x_string = "(&S)"
    x_quant = 0
    MSCommAutomacao.Output = x_string
    MsgBox "Tecle Enter para visualizar a situaçao das Bombas!", vbInformation, "Testa Automação das Bombas!"
    
    'Lê Status das bombas
    x_quant = MSCommAutomacao.InBufferCount
    x_string = MSCommAutomacao.Input
    Call CriaLogCupom(Time & " - AutomacaoLeStatusBomba: x_string=" & x_string & " x_quant=" & x_quant)
    MsgBox "x_string=" & x_string & " x_quant=" & x_quant
    
    'Fecha porta de comunicação
    MSCommAutomacao.PortOpen = False
End Sub
Private Sub AutomacaoPreparaAberturaBomba()
    lAutomacaoIlha = 1
    If MovimentoBomba.LocalizarBicoAnteriorData(g_empresa, lData, lPeriodo, lAutomacaoBicoEncerrante) Then
        lAutomacaoAbertura = MovimentoBomba.Encerrante
        lAutomacaoPeriodo = lPeriodo
        lAutomacaoData = lData
    Else
        lAutomacaoAbertura = 0
        lAutomacaoData = Date
        lAutomacaoPeriodo = 1
    End If
End Sub
Public Sub AutomacaoEnviaComandoEncerrantes(ByVal pBico As String)
    Dim i As Integer
    Dim x_quant As Integer
    Dim x_string As String
    
    'Pede encerrante
    MSCommAutomacao.InBufferCount = 0
    x_string = "(&T" & pBico & "L"
    x_quant = 0
    For i = 2 To Len(x_string)
        x_quant = x_quant + Asc(Mid(x_string, i, 1))
        If x_quant > 256 Then
            x_quant = x_quant - 256
        End If
    Next
    If Len(Hex(x_quant)) = 1 Then
        x_string = x_string & "0" & Hex(x_quant) & ")"
    Else
        x_string = x_string & Hex(x_quant) & ")"
    End If
    Call CriaLogCupom(Time & " - Movimento de Bomba: AutomacaoEnviaComandoEncerrantes: x_string=" & x_string & " Bico" & lAutomacaoBicoEncerrante)
    MSCommAutomacao.Output = x_string
End Sub
Function BuscaDataECF() As Date
    Dim xData As String
    Dim xHora As String
    BuscaDataECF = CDate("01/01/1900")
    xData = Space(6)
    xHora = Space(6)
    BemaRetorno = Bematech_FI_DataHoraImpressora(xData, xHora)
    If Len(Trim(xData)) = 0 Then
        BuscaDataECF = Date
    Else
        BuscaDataECF = CDate(Mid(xData, 1, 2) & "/" & Mid(xData, 3, 2) & "/20" & Mid(xData, 5, 2))
    End If
End Function
Private Sub BuscaNovaBomba()
    Dim i As Integer
    Dim x_data As Date
    Dim x_periodo As Integer
    Dim x_bomba As Integer
    Dim x_ilha As Integer
    
    If MovimentoBomba.LocalizarUltimo(g_empresa) Then
        txtEncerrante.SetFocus
        txtAbertura.Text = Format(MovimentoBomba.Encerrante, "###,##0.000")
        txtNumeroSubCaixa.Text = MovimentoBomba.SubCaixa
        If MovimentoBomba.SubCaixa = 999 Then
            chkSubCaixa.Value = 0
            txtNumeroSubCaixa.Enabled = False
        Else
            chkSubCaixa.Value = 1
            txtNumeroSubCaixa.Enabled = True
        End If
        x_periodo = MovimentoBomba.Periodo
        x_data = MovimentoBomba.Data
        x_bomba = MovimentoBomba.CodigoBomba
        x_ilha = MovimentoBomba.NumeroIlha
        If Bomba.LocalizarUltimo(g_empresa) = False Then
            MsgBox "Não foi possível localizar registro.", vbInformation, "Erro de Integridade"
        End If
        If x_bomba = Bomba.Codigo Then
            If Bomba.LocalizarPrimeiro(g_empresa) = False Then
                MsgBox "Não foi possível localizar registro.", vbInformation, "Erro de Integridade"
            End If
            x_bomba = Bomba.Codigo
            x_ilha = Bomba.NumeroIlha
            x_periodo = x_periodo + 1
        Else
            If Bomba.LocalizarCodigo(g_empresa, x_bomba) = False Then
                MsgBox "Não foi possível localizar registro.", vbInformation, "Erro de Integridade"
            End If
            If Bomba.LocalizarProximo Then
                x_bomba = Bomba.Codigo
                x_ilha = Bomba.NumeroIlha
            Else
                MsgBox "Não foi possível localizar registro.", vbInformation, "Erro de Integridade"
            End If
        End If
        If x_periodo > l_qtd_periodo Then
            x_data = MovimentoBomba.Data + 1
            x_periodo = 1
        End If
        txtValorLitro.Text = Format(Bomba.PrecoVenda, "###,##0.0000")
        txt_tanque.Text = Bomba.NumeroTanque
            
        For i = 0 To cbo_bomba.ListCount - 1
            cbo_bomba.ListIndex = i
            If cbo_bomba.ItemData(i) = x_ilha & Format(x_bomba, "00") Then
                Exit For
            End If
        Next
        cbo_periodo.ListIndex = x_periodo - 1
        txtData.Text = Format(x_data, "dd/mm/yyyy")
        txtAbertura.Text = Format(MovimentoBomba.EncerranteBicoAnatesDataPeriodo(g_empresa, x_data, x_bomba, x_periodo, 999), "###,##0.000")
    Else
        txtData.SetFocus
    End If
End Sub
Function ExisteRegistro(x_data As Date) As Boolean
    ExisteRegistro = False
    If MovimentoBomba.LocalizarCodigo(g_empresa, x_data, 1, lQuantidadeBico, lNumeroSubCaixa) Then
        ExisteRegistro = True
    Else
        Call CriaLogCupom(Time & " - Movimento de Bomba: Registro Inexistente: Data=" & x_data & " Ilha=" & l_ultima_ilha & " SubCaixa=" & lNumeroSubCaixa & " Bico=" & lQuantidadeBico)
    End If
End Function
Function ExisteRegistroDataPeriodo(x_data As Date, x_periodo As Integer) As Boolean
    If RetiraGString(1) = "deletar" Then
        x_periodo = 1
        ExisteRegistroDataPeriodo = False
        If Time >= CDate("06:00:00") And Time < CDate("14:00:00") Then
            x_periodo = 1
        ElseIf Time >= CDate("14:00:00") And Time < CDate("22:00:00") Then
            x_periodo = 2
        ElseIf Time >= CDate("22:00:00") And Time < CDate("23:00:00") Then
            x_periodo = 3
        ElseIf Time >= CDate("23:00:00") And Time < CDate("23:59:00") Then
            x_periodo = 4
        End If
    End If
    If MovimentoBomba.LocalizarCodigo(g_empresa, x_data, x_periodo, lQuantidadeBico, lNumeroSubCaixa) Then
        ExisteRegistroDataPeriodo = True
    End If
End Function
Private Sub Finaliza()
    Dim xCodigoFuncionario As Integer
    Dim xNomeFuncionario As String
    Call GravaAuditoria(1, Me.name, 11, "")
    If Len(g_string) > 0 Then
        If RetiraGString(1) = "DigitaEncerrantes" Then
            Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Fechado.")
            If ExisteRegistro(RetiraGString(2)) Then
                Call RegravaDataHora(RetiraGString(2))
                g_string = ""
            End If
        End If
    End If
    If fEcfInstalada = True Or UCase(g_nome_usuario) = "CUPOM FISCAL" Then
        Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Fechado.")
    End If
    If Len(g_string) > 0 Then
        If RetiraGString(1) <> "DigitaEncerrantes" Then
            xCodigoFuncionario = RetiraGString(2)
            xNomeFuncionario = RetiraGString(3)
            If RetiraGString(1) = "deletar" Then
                If ExisteRegistroDataPeriodo(RetiraGString(2), 0) Then
                    Call RegravaDataHora(RetiraGString(2))
                    GravaFechamentoCaixa
                    g_string = ""
                End If
            End If
            If RetiraGString(1) = "FechamentoCaixa" Then
                If ExisteRegistroDataPeriodo(lData, lPeriodo) Then
                    GravaFechamentoCaixa
                    g_string = ""
                End If
            End If
        End If
    End If
    If g_automacao = True Then
        If MSCommAutomacao.PortOpen = True Then
            Call RelatorioResumoVendas(RetiraGString(1))
            g_string = "cupom_complementar" & "|@|"
        End If
    End If
    flag_movimento_bomba = 0
    
    Set Bomba = Nothing
    Set Combustivel = Nothing
    Set Configuracao = Nothing
    Set ConfiguracaoDiversa = Nothing
    Set ECF = Nothing
    Set EncerranteAtual = Nothing
    Set FechamentoCaixa = Nothing
    Set LivroLMC = Nothing
    Set MovimentoAbastecimento = Nothing
    Set MovimentoBomba = Nothing
    Set MovimentoCupomFiscal = Nothing
End Sub
Private Sub GravaFechamentoCaixa()
    Dim x_cupom_inicial As Long
    Dim x_cupom_final As Long
    Dim x_hora_inicial As Date
    Dim x_hora_final As Date
    On Error GoTo ErrorGravaFechamentoCaixa
    x_cupom_inicial = 1
    If FechamentoCaixa.LocalizarUltimo(g_empresa) Then
        x_cupom_inicial = FechamentoCaixa.CupomFinal + 1
        If FechamentoCaixa.Excluir(g_empresa, CDate(txtData.Text), Val(cbo_periodo.Text)) Then
        
        End If
    End If
    x_cupom_final = 0
    If MovimentoCupomFiscal.LocalizarUltimo(g_empresa, lCodigoEcf) Then
        x_cupom_final = MovimentoCupomFiscal.NumeroCupom
    End If
    x_hora_inicial = CDate(Mid(FechamentoCaixa.HoraFinal, 1, 6) & Format(Mid(x_hora_inicial, 7, 2) + 1, "00"))
    If Val(cbo_periodo.Text) = 1 Then
        x_hora_inicial = "05:00:00"
    End If
    
    x_hora_final = "00:00:00"
    If MovimentoAbastecimento.LocalizarUltimo(g_empresa) Then
        x_hora_final = MovimentoAbastecimento.Hora
    End If
    FechamentoCaixa.Empresa = g_empresa
    FechamentoCaixa.Data = CDate(txtData.Text)
    FechamentoCaixa.Periodo = Val(cbo_periodo.Text)
    FechamentoCaixa.CupomInicial = x_cupom_inicial
    FechamentoCaixa.CupomFinal = x_cupom_final
    FechamentoCaixa.HoraInicial = x_hora_inicial
    FechamentoCaixa.HoraFinal = x_hora_final
    If Not FechamentoCaixa.Incluir Then
        MsgBox "Não foi possível incluir registro Fechamento_Caixa!", vbInformation, "Erro de Integridade!"
    End If
    Exit Sub
ErrorGravaFechamentoCaixa:
    Exit Sub
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    BuscaNovaBomba
End Sub
Private Sub IncluiSaidaBomba(pTipoCombustivel As String, pQuantidade As Currency)
    If Combustivel.LocalizarCodigo(g_empresa, pTipoCombustivel) Then
        Combustivel.QuantidadeEmEstoque = Combustivel.QuantidadeEmEstoque - pQuantidade
        If Not Combustivel.Alterar(g_empresa, pTipoCombustivel) Then
            MsgBox "Não foi possível alterar este registro.", vbInformation, "Erro de Integridade"
        End If
    Else
        MsgBox "Não foi possível localizar este registro.", vbInformation, "Erro de Integridade"
    End If
End Sub
Private Sub ExcluiSaidaBomba(pTipoCombustivel As String, pQuantidade As Currency)
    If Combustivel.LocalizarCodigo(g_empresa, pTipoCombustivel) Then
        Combustivel.QuantidadeEmEstoque = Combustivel.QuantidadeEmEstoque + pQuantidade
        If Not Combustivel.Alterar(g_empresa, pTipoCombustivel) Then
            MsgBox "Não foi possível alterar este registro.", vbInformation, "Erro de Integridade"
        End If
    Else
        MsgBox "Não foi possível localizar este registro.", vbInformation, "Erro de Integridade"
    End If
End Sub
Function PortaDaAutomacao() As Integer
    'Dim dados As String
    'Dim NumeroArquivo As Integer
    
    On Error GoTo FileError
    
    PortaDaAutomacao = 2
    PortaDaAutomacao = ReadINI("CUPOM FISCAL", "Porta da Automacao", gArquivoIni)
    
    
    'NumeroArquivo = FreeFile
    'Open "C:\VB5\SGP\CUPOM_DEMONSTRACAO.TXT" For Input As NumeroArquivo
    'If Not EOF(NumeroArquivo) Then
    '    Do Until EOF(NumeroArquivo)
    '        Line Input #NumeroArquivo, dados
    '        If Mid(dados, 1, 19) = "PORTA DA AUTOMACAO:" Then
    '            PortaDaAutomacao = Val(Mid(dados, 20, 1))
    '        End If
    '    Loop
    'End If
    'Close #NumeroArquivo
    Exit Function
FileError:
    Exit Function
End Function
Private Sub PreencheCboBomba()
    Dim nome_combustivel As String
    Dim rsBomba As New adodb.Recordset
    Dim xSQL As String
    
    xSQL = "SELECT Codigo, [Numero da Ilha], [Tipo de Combustivel] FROM Bomba WHERE Empresa = " & g_empresa & " ORDER BY Codigo"
    Set rsBomba = Conectar.RsConexao(xSQL)
    cbo_bomba.Clear
    cbo_bomba_2.Clear
    With rsBomba
        If .RecordCount > 0 Then
            Do Until .EOF
                If Combustivel.LocalizarCodigo(g_empresa, rsBomba("Tipo de Combustivel").Value) Then
                    nome_combustivel = Combustivel.Nome
                Else
                    nome_combustivel = "Não Cadastrado"
                End If
                cbo_bomba.AddItem rsBomba("Numero da Ilha").Value & Format(rsBomba("Codigo").Value, "00") & " - " & nome_combustivel
                cbo_bomba.ItemData(cbo_bomba.NewIndex) = rsBomba("Numero da Ilha").Value & Format(rsBomba("Codigo").Value, "00")
                cbo_bomba_2.AddItem rsBomba("Codigo").Value & " - " & nome_combustivel
                cbo_bomba_2.ItemData(cbo_bomba_2.NewIndex) = rsBomba("Codigo").Value
                .MoveNext
            Loop
        End If
    End With
    rsBomba.Close
    Set rsBomba = Nothing
End Sub
Private Sub cbo_bomba_2_GotFocus()
    SendMessageLong cbo_bomba_2.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_bomba_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_encerrante_inicial.SetFocus
    End If
End Sub
Private Sub cbo_bomba_2_LostFocus()
    If cbo_bomba_2.ListIndex <> -1 And IsDate(txtDataInicial.Text) Then
        txt_encerrante_inicial.Text = MovimentoBomba.AberturaBicoDataPeriodo(g_empresa, CDate(txtDataInicial.Text), Val(cbo_bomba_2.Text), 1)
    End If
End Sub
Private Sub cbo_bomba_GotFocus()
    SendMessageLong cbo_bomba.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_bomba_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtAbertura.SetFocus
    End If
End Sub
Private Sub cbo_bomba_LostFocus()
    If cbo_bomba.ListIndex <> -1 Then
        lIlha = Mid(cbo_bomba.ItemData(cbo_bomba.ListIndex), 1, 1)
        lBomba = Mid(cbo_bomba.ItemData(cbo_bomba.ListIndex), 2, 2)
        If Bomba.LocalizarCodigo(g_empresa, lBomba) Then
            txtValorLitro.Text = Format(Bomba.PrecoVenda, "###0.0000")
            txt_tanque.Text = Bomba.NumeroTanque
            txtAbertura.Text = Format(MovimentoBomba.EncerranteBicoAnatesDataPeriodo(g_empresa, CDate(txtData.Text), lBomba, Val(cbo_periodo.ItemData(cbo_periodo.ListIndex)), Val(txtNumeroSubCaixa.Text)), "###,##0.000")
        Else
            txtValorLitro = 0
        End If
    Else
        cbo_bomba.SetFocus
    End If
End Sub
Private Sub cbo_periodo_GotFocus()
    SendMessageLong cbo_periodo.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_bomba.SetFocus
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
Private Sub chkSubCaixa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If chkSubCaixa.Value = 0 Then
            cbo_bomba.SetFocus
        Else
            If txtNumeroSubCaixa.Enabled Then
                txtNumeroSubCaixa.SetFocus
            Else
                cbo_bomba.SetFocus
            End If
        End If
    End If
End Sub
Private Sub chkSubCaixa_LostFocus()
    If chkSubCaixa.Value = 0 Then
        txtNumeroSubCaixa.Text = "999"
        txtNumeroSubCaixa.Enabled = False
    Else
        txtNumeroSubCaixa.Text = "001"
        txtNumeroSubCaixa.Enabled = True
    End If
End Sub
Private Sub cmd_alterar_Click()
    Call GravaAuditoria(1, Me.name, 3, "")
    If Len(g_string) > 0 Then
        If RetiraGString(1) = "DigitaEncerrantes" Then
            Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão Alterar.")
        End If
    End If
    If fEcfInstalada = True Or UCase(g_nome_usuario) = "CUPOM FISCAL" Then
        Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão Alterar.")
    End If
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frm_dados.Enabled = True
    txtEncerrante.SetFocus
    If g_lmc = 2 Then
        txtQuantidade.MaxLength = 20
        txtQuantidade.SetFocus
    End If
End Sub
Private Sub cmd_anterior_Click()
    Call GravaAuditoria(1, Me.name, 13, "")
    If Len(g_string) > 0 Then
        If RetiraGString(1) = "DigitaEncerrantes" Then
            Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão Anterior.")
        End If
    End If
    If fEcfInstalada = True Or UCase(g_nome_usuario) = "CUPOM FISCAL" Then
        Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão Anterior.")
    End If
    If MovimentoBomba.LocalizarAnterior Then
        AtualTela
    Else
        cmd_proximo.SetFocus
        MsgBox "Início de Arquivo.", 48, "Atenção!"
    End If
End Sub
Private Sub cmd_cancelar_Click()
    Call GravaAuditoria(1, Me.name, 9, "")
    If Len(g_string) > 0 Then
        If RetiraGString(1) = "DigitaEncerrantes" Then
            Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão Cancelar.")
        End If
    End If
    If fEcfInstalada = True Or UCase(g_nome_usuario) = "CUPOM FISCAL" Then
        Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão Cancelar.")
    End If
    LimpaTela
    If MovimentoBomba.LocalizarCodigo(g_empresa, lData, lPeriodo, lBomba, lNumeroSubCaixa) Then
        AtualTela
        AtivaBotoes
        VerificaLiberacaoDigitacao
        If cmd_alterar.Enabled Then
            cmd_alterar.SetFocus
        Else
            cmd_novo.SetFocus
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
    txtData.Text = ""
    cbo_periodo.ListIndex = -1
    chkSubCaixa.Value = 0
    txtNumeroSubCaixa.Enabled = False
    txtNumeroSubCaixa.Text = ""
    cbo_bomba.ListIndex = -1
    txtValorLitro.Text = ""
    txtAbertura.Text = ""
    txtEncerrante.Text = ""
    txtQuantidade.Text = ""
    lblValorSaida.Caption = ""
    txt_tanque.Text = ""
End Sub
Private Sub cmd_cancelar2_Click()
    Call GravaAuditoria(1, Me.name, 9, "Cancelar2")
    frm_calcula_encerrante.Visible = False
    AtivaBotoes
    cmd_sair.SetFocus
End Sub
Private Sub cmd_encerrante_Click()
    DesativaBotoes
    frm_calcula_encerrante.Top = 0
    frm_calcula_encerrante.Left = 1800
    frm_calcula_encerrante.Visible = True
    frm_calcula_encerrante.ZOrder 0
    txtDataInicial.SetFocus
End Sub
Private Sub cmd_excluir_Click()
    Call GravaAuditoria(1, Me.name, 4, "")
    If Len(g_string) > 0 Then
        If RetiraGString(1) = "DigitaEncerrantes" Then
            Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão Excluir.")
        End If
    End If
    If fEcfInstalada = True Or UCase(g_nome_usuario) = "CUPOM FISCAL" Then
        Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão Excluir.")
    End If
    If IsDate(txtData.Text) Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            Call GravaAuditoria(1, Me.name, 10, "Dt:" & txtData.Text & " P:" & Val(cbo_periodo.Text) & " SubCx:" & txtNumeroSubCaixa.Text & " B:" & Val(cbo_bomba.Text) & " Abe:" & txtAbertura.Text & " Enc:" & txtEncerrante.Text)
            If Not MovimentoBomba.Excluir(g_empresa, lData, lPeriodo, lBomba, lNumeroSubCaixa) Then
                MsgBox "Não foi possível excluir este registro", vbInformation, "Erro de Integridade!"
            Else
                Call ExcluiSaidaBomba(lTipoCombustivel, lQuantidade)
            End If
            LimpaTela
            If MovimentoBomba.LocalizarUltimo(g_empresa) Then
                AtualTela
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
'    zzMudaPrecos
'    zzClonaMovimento
'    zzTransfereVendaBicos
'    zzTransfereVendaBicos
'    zzClonaMovimento2
'    zzAgrupaMovimento
'    Exit Sub
    
    Call GravaAuditoria(1, Me.name, 2, "")
    If Len(g_string) > 0 Then
        If RetiraGString(1) = "DigitaEncerrantes" Then
            Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão Novo.")
        End If
    End If
    If fEcfInstalada = True Or UCase(g_nome_usuario) = "CUPOM FISCAL" Then
        Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão Novo.")
    End If
    LimpaTela
    frm_dados.Enabled = True
    Inclui
End Sub
Private Sub cmd_novo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 7 Then 'Crtl + G
'        zzClonaMovimento2
    ElseIf KeyAscii = 19 Then 'Crtl + S
        AutomacaoLeStatusBomba
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If Len(g_string) > 0 Then
        If RetiraGString(1) = "DigitaEncerrantes" Then
            Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão OK.")
        End If
    End If
    If fEcfInstalada = True Or UCase(g_nome_usuario) = "CUPOM FISCAL" Then
        Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão OK.")
    End If
    If ValidaCampos Then
        If VerificaLiberacaoDigitacao2 Then
            AtivaBotoes
            If lOpcao = 1 Then
                Call GravaAuditoria(1, Me.name, 10, "Dt:" & txtData.Text & " P:" & Val(cbo_periodo.Text) & " SubCx:" & txtNumeroSubCaixa.Text & " B:" & Val(cbo_bomba.Text) & " Abe:" & txtAbertura.Text & " Enc:" & txtEncerrante.Text)
                AtualTabe
                If MovimentoBomba.Incluir Then
                    lData = CDate(txtData.Text)
                    lPeriodo = cbo_periodo.Text
                    lBomba = Mid(cbo_bomba.ItemData(cbo_bomba.ListIndex), 2, 2)
                    lIlha = Mid(cbo_bomba.ItemData(cbo_bomba.ListIndex), 1, 1)
                    lNumeroSubCaixa = txtNumeroSubCaixa.Text
                    lTipoCombustivel = Bomba.TipoCombustivel
                    lQuantidade = fValidaValor(txtQuantidade.Text)
                Else
                    MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
                End If
                Call IncluiSaidaBomba(lTipoCombustivel, lQuantidade)
                If Len(g_string) > 0 Then
                    If RetiraGString(1) = "DigitaEncerrantes" Then
                        If lBomba = lQuantidadeBico Then
                            cmd_sair_Click
                            g_string = "cupom_complementar" & "|@|"
                            Exit Sub
                        End If
                    End If
                End If
                LimpaTela
                Inclui
                Exit Sub
            ElseIf lOpcao = 2 Then
                Call GravaAuditoria(1, Me.name, 10, "De: Dt:" & Format(MovimentoBomba.Data, "dd/mm/yyyy") & " P:" & MovimentoBomba.Periodo & " SubCx:" & MovimentoBomba.SubCaixa & " B:" & MovimentoBomba.CodigoBomba & " Abe:" & MovimentoBomba.Abertura & " Enc:" & MovimentoBomba.Encerrante)
                Call GravaAuditoria(1, Me.name, 10, "Para: Dt:" & txtData.Text & " P:" & Val(cbo_periodo.Text) & " SubCx:" & txtNumeroSubCaixa.Text & " B:" & Val(cbo_bomba.Text) & " Abe:" & txtAbertura.Text & " Enc:" & txtEncerrante.Text)
                Call ExcluiSaidaBomba(lTipoCombustivel, lQuantidade)
                AtualTabe
                If MovimentoBomba.Alterar(g_empresa, lData, lPeriodo, lBomba, lNumeroSubCaixa) Then
                    lData = CDate(txtData.Text)
                    lPeriodo = cbo_periodo.Text
                    lBomba = Mid(cbo_bomba.ItemData(cbo_bomba.ListIndex), 2, 2)
                    lIlha = Mid(cbo_bomba.ItemData(cbo_bomba.ListIndex), 1, 1)
                    lNumeroSubCaixa = txtNumeroSubCaixa.Text
                    lTipoCombustivel = Bomba.TipoCombustivel
                    lQuantidade = fValidaValor(txtQuantidade.Text)
                Else
                    MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
                End If
                Call IncluiSaidaBomba(lTipoCombustivel, lQuantidade)
            End If
            If MovimentoBomba.LocalizarCodigo(g_empresa, lData, lPeriodo, lBomba, lNumeroSubCaixa) Then
                AtualTela
            Else
                MsgBox "Não foi possível localizar registro", vbInformation, "Erro de Integridade!"
            End If
            If lOpcao = 2 And g_lmc = 2 Then
                If Not MovimentoBomba.RecalculaEncerrante(g_empresa, lData, lBomba) Then
                    MsgBox "Erro ao recalcular encerrantes.", vbInformation, "Erro de Integridade"
                End If
                If MovimentoBomba.LocalizarCodigo(g_empresa, lData, lPeriodo, lBomba, lNumeroSubCaixa) Then
                    AtualTela
                Else
                    MsgBox "Não foi possível localizar registro", vbInformation, "Erro de Integridade!"
                End If
            End If
            lOpcao = 0
            cmd_novo.SetFocus
        End If
    End If
    Exit Sub
FileError:
    MsgBox "Erro na rotina: cmd_ok", vbInformation, "Erro não Identificado!"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(txtData.Text) Then
        MsgBox "Informe a data do movimento.", vbInformation, "Atenção!"
        txtData.SetFocus
    ElseIf CDate(txtData.Text) <> lDataECF And fEcfInstalada = True And g_nivel_acesso <> 1 And lDataECF <> "00:00:00" Then
        MsgBox "A data do movimento não pode ser diferente que " & Format(lDataECF, "dd/mm/yyyy") & ".", vbInformation, "Atenção!"
        txtData.SetFocus
    ElseIf cbo_periodo.ListIndex = -1 Then
        MsgBox "Selecione o periodo.", vbInformation, "Atenção!"
        cbo_periodo.SetFocus
    ElseIf cbo_bomba.ListIndex = -1 Then
        MsgBox "Selecione a bomba.", vbInformation, "Atenção!"
        cbo_bomba.SetFocus
    ElseIf fValidaValor4(txtValorLitro.Text) < 0 Then
        MsgBox "Informe o valor do litro.", vbInformation, "Atenção!"
        txtValorLitro.SetFocus
    ElseIf txtAbertura.Text = "" Then
        MsgBox "Informe o número de abertura.", vbInformation, "Atenção!"
        txtAbertura.SetFocus
    ElseIf txtEncerrante.Text = "" Then
        MsgBox "Informe o número do encerrante.", vbInformation, "Atenção!"
        txtEncerrante.SetFocus
    ElseIf txt_tanque.Text = "" Then
        MsgBox "Informe o número do tanque.", vbInformation, "Atenção!"
        txt_tanque.SetFocus
    ElseIf fValidaValor(txtQuantidade.Text) > lQuantidadeMaximaBico Then
        If (MsgBox("A quantidade máxima por bico é de: " & Format(lQuantidadeMaximaBico, "###,##0") & vbCrLf & "Por este motivo a quantidade está acima de uma normalidade padrão!" & vbCrLf & "Deseja realmente manter esta quantidade?", 4 + 32 + 256, "Atenção!")) = vbYes Then
            ValidaCampos = True
        Else
            MsgBox "A quantidade deve ser menor.", vbInformation, "Atenção!"
            txtEncerrante.SetFocus
        End If
    ElseIf fValidaValor(txtEncerrante.Text) < fValidaValor(txtAbertura.Text) Then
        If (MsgBox("O encerrante está menor que a abertura!" & Chr(10) & "Deseja realmente manter este encerrante menor?", 4 + 32 + 256, "Atenção!")) = 6 Then
            ValidaCampos = True
        Else
            MsgBox "O encerrante deve ser maior.", vbInformation, "Atenção!"
            txtEncerrante.SetFocus
        End If
    Else
        ValidaCampos = True
    End If
End Function
Function ValidaCampos2() As Integer
    ValidaCampos2 = False
    If Not IsDate(txtDataInicial.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        txtDataInicial.SetFocus
    ElseIf Not IsDate(txtDataFinal.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        txtDataFinal.SetFocus
    ElseIf CDate(txtDataFinal.Text) < CDate(txtDataInicial.Text) Then
        MsgBox "Data final deve ser maior que a data inicial.", vbInformation, "Atenção!"
        txtDataFinal.SetFocus
    ElseIf cbo_bomba_2.ListIndex = -1 Then
        MsgBox "Selecione o bico.", vbInformation, "Atenção!"
        cbo_bomba_2.SetFocus
    ElseIf txt_encerrante_inicial.Text = "" Then
        MsgBox "Informe o número do encerrante inicial.", vbInformation, "Atenção!"
        txt_encerrante_inicial.SetFocus
    Else
        ValidaCampos2 = True
    End If
End Function
Private Sub VerificaLiberacaoDigitacao()
    Dim x_flag As Boolean
    x_flag = True
    If g_nivel_acesso > 4 Then
        If MovimentoBomba.Empresa < g_cfg_empresa_i Or MovimentoBomba.Empresa > g_cfg_empresa_f Then
            x_flag = False
        ElseIf MovimentoBomba.Data < g_cfg_data_i Or MovimentoBomba.Data > g_cfg_data_f Then
            x_flag = False
        ElseIf MovimentoBomba.Periodo < g_cfg_periodo_i Or MovimentoBomba.Periodo > g_cfg_periodo_f Then
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
    If g_nivel_acesso <= 4 Or UCase(g_nome_usuario) = "CUPOM FISCAL" Then
        VerificaLiberacaoDigitacao2 = True
        Exit Function
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
Private Sub RegravaDataHora(x_data As Date)
    Dim x_hora As Date
    Dim x_hora_s As String
    'Dim NumeroArquivo As Integer
    Dim xQuantidadeFechamento As Integer
    Dim xHoraFechamento(1 To 8) As String
    Dim xString As String
    Dim xAlterou As Boolean
    Dim i As Integer
    'Kill "C:\VB5\SGP\DATAHORA.TXT"
    x_hora = RetiraGString(3)
    x_hora_s = x_hora
    'NumeroArquivo = FreeFile
    'Open "C:\VB5\SGP\DATAHORA.TXT" For Output As #NumeroArquivo
    
    If Configuracao.LocalizarCodigo(g_empresa) Then
        xQuantidadeFechamento = 0
        If Configuracao.HoraFechamento1 <> "00:00:00" Then
            xQuantidadeFechamento = xQuantidadeFechamento + 1
            xHoraFechamento(xQuantidadeFechamento) = Configuracao.HoraFechamento1
        End If
        If Configuracao.HoraFechamento2 <> "00:00:00" Then
            xQuantidadeFechamento = xQuantidadeFechamento + 1
            xHoraFechamento(xQuantidadeFechamento) = Configuracao.HoraFechamento2
        End If
        If Configuracao.HoraFechamento3 <> "00:00:00" Then
            xQuantidadeFechamento = xQuantidadeFechamento + 1
            xHoraFechamento(xQuantidadeFechamento) = Configuracao.HoraFechamento3
        End If
        If Configuracao.HoraFechamento4 <> "00:00:00" Then
            xQuantidadeFechamento = xQuantidadeFechamento + 1
            xHoraFechamento(xQuantidadeFechamento) = Configuracao.HoraFechamento4
        End If
        If Configuracao.HoraFechamento5 <> "00:00:00" Then
            xQuantidadeFechamento = xQuantidadeFechamento + 1
            xHoraFechamento(xQuantidadeFechamento) = Configuracao.HoraFechamento5
        End If
        If Configuracao.HoraFechamento6 <> "00:00:00" Then
            xQuantidadeFechamento = xQuantidadeFechamento + 1
            xHoraFechamento(xQuantidadeFechamento) = Configuracao.HoraFechamento6
        End If
        If Configuracao.HoraFechamento7 <> "00:00:00" Then
            xQuantidadeFechamento = xQuantidadeFechamento + 1
            xHoraFechamento(xQuantidadeFechamento) = Configuracao.HoraFechamento7
        End If
        If Configuracao.HoraFechamento8 <> "00:00:00" Then
            xQuantidadeFechamento = xQuantidadeFechamento + 1
            xHoraFechamento(xQuantidadeFechamento) = Configuracao.HoraFechamento8
        End If
        xAlterou = False
        For i = 1 To xQuantidadeFechamento
            If x_hora < CDate(xHoraFechamento(i)) Then
                x_hora_s = xHoraFechamento(i)
                xAlterou = True
                Exit For
            End If
        Next
        If xAlterou = False Then
            x_hora_s = xHoraFechamento(1)
            x_data = x_data + 1
        End If
        xString = "DATA=" & x_data & " - HORA=" & x_hora_s
        Call WriteINI("CUPOM FISCAL", "Informa Encerrante na Data", CStr(x_data), gArquivoIni)
        Call WriteINI("CUPOM FISCAL", "Informa Encerrante na Hora", x_hora_s, gArquivoIni)
        'Print #NumeroArquivo, xString
        'Close #NumeroArquivo
        Call CriaLogCupom(Time & " - Movimento de Bomba: O arquivo DataHora foi alterado Para:" & xString)
    End If
    Exit Sub
FileError:
    Call CriaLogCupom(Time & " - RegravaDataHora: Erro: " & Error & ". xString=" & xString)
    MsgBox "O arquivo temporizador DATAHORA.TXT não foi regravado!", vbCritical, "RegravaDataHora"
    Exit Sub
End Sub
Private Sub RelatorioResumoVendas(ByVal xOrigem As String)
    Dim xString As String
    Dim xString2 As String
    Dim i As Integer
    Dim i2 As Integer
    Dim xCodigoBomba As Integer
    Dim xValor As Currency
    Dim xAbertura As Currency
    Dim xEncerrante As Currency
    Dim xQuantidade As Currency
    Dim xPrecoVenda As Currency
    Dim xTotalQuantidade As Currency
    Dim xTotalBico As Currency
    Dim xImprime As Boolean
    Dim xTipoCombustivel As String
    Dim xRegistroLocalizado As Boolean
    
    xValor = 0
    xTotalQuantidade = 0
    lQtdA = 0
    lQtdD = 0
    lQtdG = 0
    lTotalA = 0
    lTotalD = 0
    lTotalG = 0
    
    
    xImprime = False
    If xOrigem = "Automacao" Then
        If (MsgBox("Deseja Imprimir o resumo de venda?", vbQuestion + vbYesNo + vbDefaultButton2, "Resumo de Venda!")) = vbYes Then
            xImprime = True
        End If
        MSCommAutomacao.PortOpen = False
    ElseIf xOrigem = "ConsultaVenda" Then
        lAutomacaoData = RetiraGString(2)
        lAutomacaoPeriodo = RetiraGString(3)
    End If
    xString = ""
    xString = xString + "BC   ABERTURA  ENCERRANTE     LITROS       VALOR"
    If xImprime = False Then
        lNomeArquivo = BioCriaImprime
        'seleciona medidas para centímetros
        BioImprime "@@Printer.ScaleMode = 7"
        BioImprime "@@Printer.PaperSize = 1"
        BioImprime "@@Printer.FontName = Draft 10cpi"
        BioImprime "@@Printer.FontName = Draft 10cpi"
        'teste para imprimir letra correta
        BioImprime "@@Printer.FontBold = False"
        BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
        xString2 = "+------------------------------------------------------------------+"
        BioImprime "@Printer.Print " & xString2
        xString2 = "| RELATORIO DE RESUMO DE VENDAS DE COMBUSTIVEIS                    |"
        BioImprime "@Printer.Print " & xString2
        xString2 = "| REFERENTE A.: __/__/____    DO PERIODO.: _                       |"
        Mid(xString2, 17, 10) = Format(lAutomacaoData, "dd/mm/yyyy")
        Mid(xString2, 44, 1) = lAutomacaoPeriodo
        BioImprime "@Printer.Print " & xString2
        xString2 = "+------+--------------+--------------+--------------+--------------+"
        BioImprime "@Printer.Print " & xString2
        xString2 = "| BICO |   ABERTURA   |  ENCERRANTE  | QTD.  LITROS | VALOR  VENDA |"
        BioImprime "@Printer.Print " & xString2
        xString2 = "+------+--------------+--------------+--------------+--------------+"
        BioImprime "@Printer.Print " & xString2
    End If
    For i = 1 To lQuantidadeBico
        '                     1         2         3         4       4
        '            123456789012345678901234567890123456789012345678
        'xString = " ) 123.123,12  123.123,12  23,123.12  123.123.12"
        xString2 = " )       0,00        0,00       0,00        0,00"
        Mid(xString, 1, 1) = i
        xRegistroLocalizado = False
        If xOrigem = "Automacao" Then
            If MovimentoBomba.LocalizarCodigo(g_empresa, lAutomacaoData, lAutomacaoPeriodo, i, 999) Then
                xCodigoBomba = MovimentoBomba.CodigoBomba
                xTipoCombustivel = Trim(MovimentoBomba.TipoCombustivel)
                xAbertura = MovimentoBomba.Abertura
                xEncerrante = MovimentoBomba.Encerrante
                xQuantidade = MovimentoBomba.QuantidadeSaida
                xPrecoVenda = MovimentoBomba.PrecoVenda
                xRegistroLocalizado = True
            End If
        ElseIf xOrigem = "ConsultaVenda" Then
            'If MovimentoBomba.LocalizarCodigo(g_empresa, lAutomacaoData, lAutomacaoPeriodo, i, 999) Then
            If MovimentoBomba.LocalizarBicoAnteriorData(g_empresa, lAutomacaoData, lAutomacaoPeriodo, i) Then
                xCodigoBomba = MovimentoBomba.CodigoBomba
                xTipoCombustivel = Trim(MovimentoBomba.TipoCombustivel)
                xAbertura = MovimentoBomba.Encerrante
                If Bomba.LocalizarCodigo(g_empresa, MovimentoBomba.CodigoBomba) Then
                    xPrecoVenda = Bomba.PrecoVenda
                Else
                    xPrecoVenda = 0
                End If
                If EncerranteAtual.LocalizarCodigo(g_empresa, MovimentoBomba.CodigoBomba) Then
                    xEncerrante = EncerranteAtual.Encerrante
                Else
                    xEncerrante = 0
                End If
                xQuantidade = xEncerrante - xAbertura
                xRegistroLocalizado = True
            End If
        End If
        If xRegistroLocalizado Then
            i2 = Len(Format(xAbertura, "###,##0.000"))
            Mid(xString2, 4 + 10 - i2, i2) = Format(xAbertura, "###,##0.000")
            i2 = Len(Format(xEncerrante, "###,##0.000"))
            Mid(xString2, 16 + 10 - i2, i2) = Format(xEncerrante, "###,##0.000")
            i2 = Len(Format(xQuantidade, "##,##0.000"))
            Mid(xString2, 28 + 9 - i2, i2) = Format(xQuantidade, "##,##0.000")
            
            xTotalBico = Format(xQuantidade * xPrecoVenda, "###,##0.00")
            i2 = Len(Format(xTotalBico, "###,##0.00"))
            Mid(xString2, 39 + 10 - i2, i2) = Format(xTotalBico, "###,##0.00")
            xValor = xValor + xTotalBico
            xTotalQuantidade = xTotalQuantidade + xQuantidade
            Select Case xTipoCombustivel
                Case "A"
                    lQtdA = lQtdA + xQuantidade
                    lTotalA = lTotalA + xTotalBico
                Case "D"
                    lQtdD = lQtdD + xQuantidade
                    lTotalD = lTotalD + xTotalBico
                Case "G"
                    lQtdG = lQtdG + xQuantidade
                    lTotalG = lTotalG + xTotalBico
            End Select
            xString = xString & xString2
            If xImprime = False Then
                '           123456789
                xString2 = "|      |              |              |              |              |"
                Mid(xString2, 4, 2) = Format(xCodigoBomba, "00")
                i2 = Len(Format(xAbertura, "###,##0.000"))
                Mid(xString2, 12 + 10 - i2, i2) = Format(xAbertura, "###,##0.000")
                i2 = Len(Format(xEncerrante, "###,##0.00"))
                Mid(xString2, 26 + 10 - i2, i2) = Format(xEncerrante, "###,##0.000")
                i2 = Len(Format(xQuantidade, "##,##0.00"))
                Mid(xString2, 43 + 9 - i2, i2) = Format(xQuantidade, "##,##0.000")
                i2 = Len(Format(xTotalBico, "###,##0.00"))
                Mid(xString2, 57 + 10 - i2, i2) = Format(xTotalBico, "###,##0.00")
                BioImprime "@Printer.Print " & xString2
            End If
        End If
    Next
    xString2 = "          TOTAL DO CAIXA:       0,00        0,00"
    i2 = Len(Format(xTotalQuantidade, "###,##0.000"))
    Mid(xString2, 27 + 10 - i2, i2) = Format(xTotalQuantidade, "###,##0.000")
    i2 = Len(Format(xValor, "###,##0.00"))
    Mid(xString2, 39 + 10 - i2, i2) = Format(xValor, "###,##0.00")
    xString = xString & xString2
    If xImprime = True Then
        'Abertura do Relatório Gerencial
        BemaRetorno = Bematech_FI_RelatorioGerencial(xString)
        'Fechamento de Relatório Gerencial
        BemaRetorno = Bematech_FI_FechaRelatorioGerencial
    Else
        
        xString2 = "+------+--------------+--------------+--------------+--------------+"
        BioImprime "@Printer.Print " & xString2
        xString2 = "|                       ALCOOL       |              |              |"
        i2 = Len(Format(lQtdA, "###,##0.000"))
        Mid(xString2, 43 + 10 - i2, i2) = Format(lQtdA, "###,##0.000")
        i2 = Len(Format(lTotalA, "###,##0.00"))
        Mid(xString2, 57 + 10 - i2, i2) = Format(lTotalA, "###,##0.00")
        BioImprime "@Printer.Print " & xString2
        xString2 = "|                       DIESEL       |              |              |"
        i2 = Len(Format(lQtdD, "###,##0.000"))
        Mid(xString2, 43 + 10 - i2, i2) = Format(lQtdD, "###,##0.000")
        i2 = Len(Format(lTotalD, "###,##0.00"))
        Mid(xString2, 57 + 10 - i2, i2) = Format(lTotalD, "###,##0.00")
        BioImprime "@Printer.Print " & xString2
        xString2 = "|                       GASOLINA     |              |              |"
        i2 = Len(Format(lQtdG, "###,##0.000"))
        Mid(xString2, 43 + 10 - i2, i2) = Format(lQtdG, "###,##0.000")
        i2 = Len(Format(lTotalG, "###,##0.00"))
        Mid(xString2, 57 + 10 - i2, i2) = Format(lTotalG, "###,##0.00")
        BioImprime "@Printer.Print " & xString2
        xString2 = "+------------------------------------+--------------+--------------+"
        BioImprime "@Printer.Print " & xString2
        xString2 = "|                       TOTAL        |              |              |"
        i2 = Len(Format(xTotalQuantidade, "###,##0.000"))
        Mid(xString2, 43 + 10 - i2, i2) = Format(xTotalQuantidade, "###,##0.000")
        i2 = Len(Format(xValor, "###,##0.00"))
        Mid(xString2, 57 + 10 - i2, i2) = Format(xValor, "###,##0.00")
        BioImprime "@Printer.Print " & xString2
        xString2 = "+--- Cerrado Informatica. -----------+--------------+--------------+"
        BioImprime "@Printer.Print " & xString2
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = "0" & lNomeArquivo & "|@|Resumo de Venda da Automaçao|@|"
        frm_preview.Show 1
        g_string = ""
    End If
    If xOrigem = "ConsultaVenda" Then
        cmd_sair_Click
    End If
End Sub
Private Sub TransfereDadosLMC()
    On Error GoTo FileError
    
    Dim xData As Date
    Dim xMensagem As String
    
    xMensagem = "para o L.M.C.!"
    
    'Busca ultima data com movimento
    If MovimentoBomba.LocalizarUltimo(g_empresa) Then
        lUltimaData = MovimentoBomba.Data
    Else
        lUltimaData = CDate("01/01/1900")
    End If
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será transferido o movimento de bomba apartir da data " & lUltimaData & "." & Chr(10) & Chr(10) & "Deseja realmente fazer esta transferência?", vbYesNo + 256, "Transfere o Movimento de Bomba " & xMensagem)) = 7 Then
        Exit Sub
    End If
    Call GravaAuditoria(1, Me.name, 10, "Empresa:" & g_empresa & " A Partir de:" & lUltimaData)
    
    'Exclui Movimento da Última Data
    If (MsgBox("É aconselhavel excluir o movimento de bomba da data " & lUltimaData & "." & Chr(10) & Chr(10) & "Deseja excluir automaticamente?", vbYesNo + 256, "Exclusão de Registros")) = vbYes Then
        If Not MovimentoBomba.ExcluirData(g_empresa, lUltimaData) Then
            MsgBox "Não foi possível excluir registros", vbInformation, "Erro de Verificação"
        End If
    End If
    
    'Transfere Dados para o LMC
    If MovimentoBomba.TransfereDados(g_empresa, lUltimaData, "Movimento_Bomba") Then
        If Not MovimentoBomba.RecalculaEncerrante(g_empresa, lUltimaData, 0) Then
            MsgBox "Erro ao recalcular encerrantes.", vbInformation, "Erro de Integridade"
        End If
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com o movimento de bomba transferido " & xMensagem, vbInformation, "Transferência Concluida!"
    Else
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Não tem movimento de bomba à ser transferido " & xMensagem, vbInformation, "Transferência Não Concluida!"
    End If
    
    'Acumula Períodos
    MsgBox "O sistema irá acumular os períodos!" & Chr(10) & "Isto pode demorar alguns minutos.", vbInformation, "Transferência para o LMC"
    xData = MovimentoBomba.LocalizarPDPM1(g_empresa)
    Do Until xData = "00:00:00"
        If Not MovimentoBomba.AgrupaPeriodoData(g_empresa, xData) Then
            MsgBox "Erro ao acumular períodos!", vbInformation, "Erro de Integridade!"
            Exit Sub
        End If
        xData = MovimentoBomba.LocalizarPDPM1(g_empresa)
    Loop
    MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com os períodos acumulados.", vbInformation, "Operação Concluida!"
    Exit Sub
    
FileError:
    MsgBox "Erro na rotina: TransfereDadosLMC", vbInformation, "Erro não Identificado!"
    Resume Next
End Sub
Private Sub TransfereDadosCaixa()
    On Error GoTo FileError
    
    Dim xData As Date
    Dim xMensagem As String
    
    xMensagem = "da Pista para o do Caixa!"
    
    'Busca ultima data com movimento
    If MovimentoBomba.LocalizarUltimo(g_empresa) Then
        lUltimaData = MovimentoBomba.Data
    Else
        lUltimaData = CDate("01/01/1900")
    End If
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será transferido o movimento de bomba apartir da data " & lUltimaData & "." & Chr(10) & Chr(10) & "Deseja realmente fazer esta transferência?", vbYesNo + 256, "Transfere o Movimento de Bomba " & xMensagem)) = 7 Then
        Exit Sub
    End If
    Call GravaAuditoria(1, Me.name, 10, "Empresa:" & g_empresa & " A Partir de:" & lUltimaData)
    
    'Exclui Movimento da Última Data
    If (MsgBox("É aconselhavel excluir o movimento de bomba da data " & lUltimaData & "." & Chr(10) & Chr(10) & "Deseja excluir automaticamente?", vbYesNo + 256, "Exclusão de Registros")) = vbYes Then
        If Not MovimentoBomba.ExcluirData(g_empresa, lUltimaData) Then
            MsgBox "Não foi possível excluir registros", vbInformation, "Erro de Verificação"
        End If
    End If
    
    'Transfere Dados para o Caixa
    xMensagem = "para o Caixa."
    If MovimentoBomba.TransfereDados(g_empresa, lUltimaData, "Movimento_Bomba_Cupom") Then
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com o movimento de bomba transferido " & xMensagem, vbInformation, "Transferência Concluida!"
    Else
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Não tem movimento de bomba à ser transferido " & xMensagem, vbInformation, "Transferência Não Concluida!"
    End If
    Exit Sub
    
FileError:
    MsgBox "Erro na rotina: TransfereDadosCaixa", vbInformation, "Erro não Identificado!"
    Resume Next
End Sub
Private Sub cmd_ok2_Click()
    If ValidaCampos2 Then
        If (MsgBox("Deseja realmente calcular os encerrantes?", 4 + 32 + 256, "Calcula Encerrante!")) = 6 Then
            If MovimentoBomba.RecalculaEncerrante2(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_bomba_2.Text), fValidaValor(txt_encerrante_inicial.Text)) Then
                MsgBox "Processamento Concluido com Sucesso!", vbOKOnly, "Calcula Encerrante"
            Else
                MsgBox "Erro no processamento!", vbCritical, "Erro de Processamento"
            End If
            frm_calcula_encerrante.Visible = False
            AtivaBotoes
            cmd_sair.SetFocus
        Else
            frm_calcula_encerrante.Visible = False
            AtivaBotoes
            cmd_sair.SetFocus
        End If
    End If
End Sub
Private Sub cmd_ok2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 1 Then
        If (MsgBox("Deseja realmente AUMENTAR as vendas?", 4 + 32 + 256, "AUMENTA Vendas!")) = 6 Then
            If MovimentoBomba.AlteraQuantidadeDatas(g_empresa, "+", CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(Me.cbo_bomba_2.Text), fValidaValor(txt_encerrante_inicial.Text)) Then
                MsgBox "Processamento Concluído!"
            Else
                MsgBox "Erro no Processamento!"
            End If
        End If
    ElseIf KeyAscii = 4 Then
        If (MsgBox("Deseja realmente DIMINUIR as vendas?", 4 + 32 + 256, "DIMINUI Vendas!")) = 6 Then
            If MovimentoBomba.AlteraQuantidadeDatas(g_empresa, "-", CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(Me.cbo_bomba_2.Text), fValidaValor(txt_encerrante_inicial.Text)) Then
                MsgBox "Processamento Concluído!"
            Else
                MsgBox "Erro no Processamento!"
            End If
        End If
    End If
End Sub
Private Sub cmd_pesquisa_Click()
    Dim xString As String
    Call GravaAuditoria(1, Me.name, 5, "")
    xString = ""
    If Len(g_string) > 0 Then
        xString = g_string
        If RetiraGString(1) = "DigitaEncerrantes" Then
            Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão Pesquisa.")
        End If
    End If
    If fEcfInstalada = True Or UCase(g_nome_usuario) = "CUPOM FISCAL" Then
        Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão Pesquisa.")
    End If
    consulta_movimento_bomba.Show 1
    If Len(g_string) > 0 Then
        lData = RetiraGString(1)
        lPeriodo = RetiraGString(2)
        lIlha = RetiraGString(3)
        lBomba = RetiraGString(4)
        lNumeroSubCaixa = RetiraGString(5)
        If MovimentoBomba.LocalizarCodigo(g_empresa, lData, lPeriodo, lBomba, lNumeroSubCaixa) Then
            AtualTela
        Else
            MsgBox "Não foi possível localizar registro", vbInformation, "Erro de Integridade!"
        End If
        g_string = ""
    End If
    If xString <> "" Then
        g_string = xString
    End If
End Sub
Private Sub cmd_primeiro_Click()
    Call GravaAuditoria(1, Me.name, 12, "")
    If Len(g_string) > 0 Then
        If RetiraGString(1) = "DigitaEncerrantes" Then
            Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão Primeiro.")
        End If
    End If
    If fEcfInstalada = True Or UCase(g_nome_usuario) = "CUPOM FISCAL" Then
        Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão Primeiro.")
    End If
    If MovimentoBomba.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    Call GravaAuditoria(1, Me.name, 14, "")
    If Len(g_string) > 0 Then
        If RetiraGString(1) = "DigitaEncerrantes" Then
            Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão Próximo.")
        End If
    End If
    If fEcfInstalada = True Or UCase(g_nome_usuario) = "CUPOM FISCAL" Then
        Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão Próximo.")
    End If
    If MovimentoBomba.LocalizarProximo Then
        AtualTela
    Else
        cmd_anterior.SetFocus
        MsgBox "Fim de Arquivo.", 48, "Atenção!"
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_transfere_dados_lmc_Click()
    If g_nome_usuario = "L.M.C." Then
        Call GravaAuditoria(1, Me.name, 23, "Transferencia Para LMC")
        TransfereDadosLMC
    Else
    Call GravaAuditoria(1, Me.name, 23, "Transferencia Para Caixa")
        TransfereDadosCaixa
    End If
    cmd_cancelar_Click
    If MovimentoBomba.LocalizarUltimo(g_empresa) Then
        AtualTela
        AtivaBotoes
        VerificaLiberacaoDigitacao
    Else
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
    End If
End Sub
Private Sub cmd_ultimo_Click()
    Call GravaAuditoria(1, Me.name, 15, "")
    If Len(g_string) > 0 Then
        If RetiraGString(1) = "DigitaEncerrantes" Then
            Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão Ultimo.")
        End If
    End If
    If fEcfInstalada = True Or UCase(g_nome_usuario) = "CUPOM FISCAL" Then
        Call CriaLogCupom(Time & " - Movimento de Bomba: Foi Acionado o Botão Ultimo.")
    End If
    If MovimentoBomba.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Len(g_string) > 0 Then
        If RetiraGString(1) = "DigitaEncerrantes" Then
            Call CriaLogCupom(Time & " - Movimento de Bomba: Abriu.")
        End If
    End If
    If fEcfInstalada = True Or UCase(g_nome_usuario) = "CUPOM FISCAL" Then
        Call CriaLogCupom(Time & " - Movimento de Bomba: Abriu.")
    End If
    If flag_movimento_bomba = 0 Then
        DesativaBotoes
        AtualizaConstantes
        If MovimentoBomba.LocalizarUltimo(g_empresa) Then
            AtualTela
            AtivaBotoes
            VerificaLiberacaoDigitacao
            Call VerificaLiberacaoLMC(lTipoCombustivel, lData)
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        If cmd_novo.Enabled Then
            cmd_novo.SetFocus
        End If
        If g_automacao = True Then
            If RetiraGString(1) = "Automacao" Then
                '----- Porta
                lAutomacaoPorta = PortaDaAutomacao
                '----- Velocidade
                lAutomacaoVelocidade = "9600,n,8,1"
                '----- DTR
                lAutomacaoDtr = True
                'ldtr = False
                '----- RTS
                lAutomacaoRts = True
                'lRts = False
                '----- EOF
                AutomacaoAbrePorta
                lAutomacaoFlag = 0
                lAutomacaoBicoEncerrante = 1
                lTentativaLeitura = 0
                TimerAutomacao.Enabled = True
            ElseIf RetiraGString(1) = "ConsultaVenda" Then
                Call RelatorioResumoVendas(RetiraGString(1))
            End If
        End If
    Else
        flag_movimento_bomba = 0
    End If
End Sub
Private Sub Form_Deactivate()
    flag_movimento_bomba = 1
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
    
    If g_nome_usuario = "L.M.C." Then
        cmd_transfere_dados_lmc.Caption = "&Transfere p/ LMC"
        MovimentoBomba.NomeTabela = "Movimento_Bomba_LMC"
        Me.Caption = Me.Caption & " - LMC"
    ElseIf UCase(g_nome_usuario) = "CUPOM FISCAL" Or fEcfInstalada Then
        lDataECF = BuscaDataECF
        MovimentoBomba.NomeTabela = "Movimento_Bomba_Cupom"
        Me.Caption = Me.Caption & " - ECF"
    Else
        cmd_transfere_dados_lmc.Caption = "&Transfere p/Caixa"
        MovimentoBomba.NomeTabela = "Movimento_Bomba"
    End If
    PreencheCboBomba
    PreencheCboPeriodo
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub txtAbertura_GotFocus()
    'If g_nivel_acesso > 4 And Configuracao.AlteraAberturaBomba = False Then
    If g_nivel_acesso > lNivelAcessoRestringeEditarEncerranteAbertura Then
        txtEncerrante.SetFocus
    Else
        txtAbertura.SelStart = 0
        txtAbertura.SelLength = Len(txtAbertura.Text)
    End If
End Sub
Private Sub txtAbertura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtEncerrante.SetFocus
    End If
End Sub
Private Sub txtAbertura_LostFocus()
    txtAbertura.Text = Format(txtAbertura.Text, "###,##0.000")
End Sub
Private Sub txtEncerrante_GotFocus()
    If Len(txtEncerrante.Text) = 0 Then
        txtEncerrante.Text = txtAbertura.Text
    End If
    txtEncerrante.SelStart = 0
    txtEncerrante.SelLength = Len(txtEncerrante.Text)
End Sub
Private Sub txtEncerrante_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    ElseIf KeyAscii = 18 Then
        KeyAscii = 0
        txtEncerrante.Text = txtAbertura.Text
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txtEncerrante_LostFocus()
    If txtEncerrante.Text <> "" Then
        txtEncerrante.Text = Format(txtEncerrante.Text, "###,##0.000")
        If g_empresa = 44 Then
            If fValidaValor(txtEncerrante.Text) < fValidaValor(txtAbertura.Text) Then
                txtQuantidade.Text = Format(fValidaValor(txtEncerrante.Text) + 10000 - fValidaValor(txtAbertura.Text), "###,##0.000")
            Else
                txtQuantidade.Text = Format(fValidaValor(txtEncerrante.Text) - fValidaValor(txtAbertura.Text), "###,##0.000")
            End If
        Else
            If fValidaValor(txtEncerrante.Text) < fValidaValor(txtAbertura.Text) Then
                txtQuantidade.Text = Format(fValidaValor(txtEncerrante.Text) + 1000000 - fValidaValor(txtAbertura.Text), "###,##0.000")
            Else
                txtQuantidade.Text = Format(fValidaValor(txtEncerrante.Text) - fValidaValor(txtAbertura.Text), "###,##0.000")
            End If
        End If
        'lblValorSaida.Caption = Format(fValidaValor(txtquantidade.Text) * Bomba.PrecoVenda, "###,##0.00")
        lblValorSaida.Caption = Format(fValidaValor(txtQuantidade.Text) * fValidaValor(txtValorLitro.Text), "###,##0.00")
    End If
End Sub
Private Sub txtQuantidade_GotFocus()
    If g_lmc <> 2 Then
        cmd_ok.SetFocus
    Else
        txtQuantidade.SelStart = 0
        txtQuantidade.SelLength = Len(txtQuantidade.Text)
    End If
End Sub
Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txtQuantidade_LostFocus()
    If fVerificaSomaSubtracao(txtQuantidade.Text) Then
        txtQuantidade.Text = fCalculaValores(txtQuantidade.Text)
    End If
    If txtQuantidade.Text <> "" Then
        txtQuantidade.Text = Format(txtQuantidade.Text, "###,##0.000")
        txtEncerrante.Text = Format(fValidaValor(txtAbertura.Text) + fValidaValor(txtQuantidade.Text), "###,##0.000")
        lblValorSaida = Format(fValidaValor(txtQuantidade.Text) * Bomba.PrecoVenda, "###,##0.00")
    End If
End Sub
Private Sub TimerAutomacao_Timer()
    'Pede Encerrante - lAutomacaoFlag = 1 e 2
    If lAutomacaoFlag = 0 Then
        Call AutomacaoEnviaComandoEncerrantes(ConverteBico(CStr(Format(lAutomacaoBicoEncerrante, "00"))))
        lAutomacaoFlag = 1
    ElseIf lAutomacaoFlag = 1 Then
        Call AutomacaoBuscaEncerrantes(ConverteBico(CStr(Format(lAutomacaoBicoEncerrante, "00"))))
        lAutomacaoFlag = 0
        If lTentativaLeitura = 6 Then
            lTentativaLeitura = 0
            lAutomacaoBicoEncerrante = lAutomacaoBicoEncerrante + 1
            If lAutomacaoBicoEncerrante > lQuantidadeBico Then
                lAutomacaoBicoEncerrante = 0
                TimerAutomacao.Enabled = False
                GravaFechamentoCaixa
                'MsgBox "Leitura de encerrantes efetuada!", vbInformation, "Leitura OK!"
                cmd_sair_Click
            End If
        End If
    End If
End Sub
Private Sub txt_encerrante_inicial_GotFocus()
    txt_encerrante_inicial.SelStart = 0
    txt_encerrante_inicial.SelLength = Len(txt_encerrante_inicial.Text)
End Sub
Private Sub txt_encerrante_inicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok2.SetFocus
    End If
End Sub
Private Sub txt_encerrante_inicial_LostFocus()
    If txt_encerrante_inicial.Text <> "" Then
        txt_encerrante_inicial.Text = Format(txt_encerrante_inicial.Text, "###,##0.000")
    End If
End Sub
Private Sub txt_tanque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtValorLitro_GotFocus()
    txtValorLitro.SelStart = 0
    txtValorLitro.SelLength = Len(txtValorLitro.Text)
End Sub
Private Sub txtValorLitro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtAbertura.SetFocus
    End If
End Sub
Private Sub txtValorLitro_LostFocus()
    txtValorLitro.Text = Format(txtValorLitro.Text, "###,##0.0000")
End Sub
Private Sub txtData_GotFocus()
    If txtData.Text = "" And lOpcao = 1 Then
        txtData.Text = Format(g_data_def, "dd/mm/yyyy")
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
Private Sub txtDataFinal_GotFocus()
    If txtDataFinal.Text = "" Then
        txtDataFinal.Text = Format(g_data_def, "dd/mm/yyyy")
    End If
    txtDataFinal.Text = fDesmascaraData(txtDataFinal.Text)
    txtDataFinal.SelStart = 0
    txtDataFinal.SelLength = 4
    txtDataFinal.MaxLength = 8
End Sub
Private Sub txtDataFinal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_bomba_2.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtDataFinal_LostFocus()
    txtDataFinal.MaxLength = 10
    txtDataFinal.Text = fMascaraData(txtDataFinal.Text)
End Sub
Private Sub txtDataInicial_GotFocus()
    If txtDataInicial.Text = "" Then
        txtDataInicial.Text = Format(g_data_def, "dd/mm/yyyy")
    End If
    txtDataInicial.Text = fDesmascaraData(txtDataInicial.Text)
    txtDataInicial.SelStart = 0
    txtDataInicial.SelLength = 4
    txtDataInicial.MaxLength = 8
End Sub
Private Sub txtDataInicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtDataFinal.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtDataInicial_LostFocus()
    txtDataInicial.MaxLength = 10
    txtDataInicial.Text = fMascaraData(txtDataInicial.Text)
    If IsDate(txtDataInicial.Text) Then
        txtDataFinal.Text = txtDataInicial.Text
    End If
End Sub
Private Sub txtNumeroSubCaixa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_bomba.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub


Private Sub zzAgrupaMovimento()
    Dim rsMovBomba As New adodb.Recordset
    Dim xSQL As String
    Dim xCodigoLocalizar As Integer
    Dim xCodigoGravar As Integer
    
    xCodigoLocalizar = 5
    xCodigoGravar = 1
    
    
    Exit Sub
    xSQL = "SELECT Data, Periodo, [Codigo da Bomba], SubCaixa, [Quantidade da Saida],"
    xSQL = xSQL & "       [Tipo de Combustivel], [Preco de Custo], [Preco de Venda]"
    xSQL = xSQL & "  FROM Movimento_Bomba_LMC"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & "   AND [Codigo da Bomba] = " & xCodigoLocalizar
    xSQL = xSQL & "   AND Data >= " & preparaData(CDate("01/01/2008"))
    xSQL = xSQL & "   AND Data <= " & preparaData(CDate("07/01/2008"))
    xSQL = xSQL & " ORDER BY Data"
    Set rsMovBomba = Conectar.RsConexao(xSQL)
    If rsMovBomba.RecordCount > 0 Then
        Do Until rsMovBomba.EOF
            'Adiciona quantidade
            If MovimentoBomba.LocalizarCodigo(g_empresa, rsMovBomba("Data").Value, rsMovBomba("Periodo").Value, xCodigoGravar, rsMovBomba("SubCaixa").Value) Then
                MovimentoBomba.QuantidadeSaida = MovimentoBomba.QuantidadeSaida + rsMovBomba("Quantidade da Saida").Value
                MovimentoBomba.TipoCombustivel = rsMovBomba("Tipo de Combustivel").Value
                MovimentoBomba.PrecoCusto = rsMovBomba("Preco de Custo").Value
                MovimentoBomba.PrecoVenda = rsMovBomba("Preco de Venda").Value
                If MovimentoBomba.Alterar(g_empresa, rsMovBomba("Data").Value, rsMovBomba("Periodo").Value, xCodigoGravar, rsMovBomba("SubCaixa").Value) = False Then
                    MsgBox "Não foi possível alterar"
                End If
            Else
                MsgBox "Movimento inexistente"
            End If
            'Zera Quantidade
            If MovimentoBomba.LocalizarCodigo(g_empresa, rsMovBomba("Data").Value, rsMovBomba("Periodo").Value, xCodigoLocalizar, rsMovBomba("SubCaixa").Value) Then
                MovimentoBomba.QuantidadeSaida = 0
                If MovimentoBomba.Alterar(g_empresa, rsMovBomba("Data").Value, rsMovBomba("Periodo").Value, xCodigoLocalizar, rsMovBomba("SubCaixa").Value) = False Then
                    MsgBox "Não foi possível alterar"
                End If
            Else
                MsgBox "Movimento inexistente"
            End If
            rsMovBomba.MoveNext
        Loop
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
End Sub
Private Sub zzClonaMovimento()
    Dim rsMovBomba As New adodb.Recordset
    Dim xSQL As String
    Dim xCodigoLocalizar As Integer
    Dim xCodigoGravar As Integer
    Dim xTipoCombustivel As String
    Dim xNumeroTanque As Integer
    
    xCodigoLocalizar = 1
    xCodigoGravar = 1
    xTipoCombustivel = "G "
    xNumeroTanque = 1
    
    
    Exit Sub
    xSQL = "SELECT Data, Periodo, [Codigo da Bomba], SubCaixa"
    xSQL = xSQL & "  FROM Movimento_Bomba_LMC"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    'xSQL = xSQL & "   AND [Codigo da Bomba] = " & xCodigoLocalizar
    xSQL = xSQL & "   AND Data >= " & preparaData(CDate("01/01/2005"))
    xSQL = xSQL & "   AND Data <= " & preparaData(CDate("31/12/2005"))
    xSQL = xSQL & " ORDER BY Data"
    Set rsMovBomba = Conectar.RsConexao(xSQL)
    If rsMovBomba.RecordCount > 0 Then
        Do Until rsMovBomba.EOF
            If MovimentoBomba.LocalizarCodigo(g_empresa, rsMovBomba("Data").Value, rsMovBomba("Periodo").Value, rsMovBomba("Codigo da Bomba").Value, rsMovBomba("SubCaixa").Value) Then
                MovimentoBomba.Data = CDate(Day(MovimentoBomba.Data) & "/" & Month(MovimentoBomba.Data) & "/" & 2007)
                'MovimentoBomba.Data = CDate(Day(MovimentoBomba.Data) & "/11" & "/" & 2005)
'                MovimentoBomba.CodigoBomba = xCodigoGravar
'                MovimentoBomba.NumeroTanque = xNumeroTanque
'                MovimentoBomba.TipoCombustivel = xTipoCombustivel
                If MovimentoBomba.Incluir = False Then
                    MsgBox "Não foi possível incluir"
                End If
            Else
                MsgBox "Movimento inexistente"
            End If
            rsMovBomba.MoveNext
        Loop
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
End Sub
Private Sub zzClonaMovimento2()
    Dim xBico As Integer
    Dim xData As Date
    
    Exit Sub
    For xBico = 1 To 14
        If MovimentoBomba.LocalizarCodigo(g_empresa, CDate("28/11/2009"), 1, xBico, 999) Then
            For xData = CDate("26/06/2006") To CDate("27/11/2009")
                MovimentoBomba.Data = xData
                MovimentoBomba.Encerrante = MovimentoBomba.Abertura
                MovimentoBomba.QuantidadeSaida = 0
                If Not MovimentoBomba.Incluir Then
                    MsgBox "Não foi possível incluir"
                End If
            Next
        Else
            MsgBox "Movimento inexistente"
        End If
    Next
End Sub
Private Sub zzMudaPrecos()
    Dim rsMovBomba As New adodb.Recordset
    Dim xSQL As String
    Dim xPrecoA As Currency
    Dim xPrecoD As Currency
    Dim xPrecoG As Currency
    Dim xPrecoGA As Currency
    Dim xData As Date
    
    Exit Sub
    xSQL = "SELECT *"
    xSQL = xSQL & "  FROM Movimento_Bomba_LMC"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & "   AND Data <= " & preparaData(CDate("30/06/2005"))
    xSQL = xSQL & " ORDER BY Data, [Codigo da Bomba]"
    Set rsMovBomba = Conectar.RsConexao(xSQL)
    If rsMovBomba.RecordCount > 0 Then
        xData = rsMovBomba("Data").Value
        Do Until rsMovBomba.EOF
            If xData <> rsMovBomba("Data").Value Then
                If Not MovimentoBomba.AlteraPrecoVendaBomba(g_empresa, xData, 1, xPrecoGA) Then
                    MsgBox "Erro Bico=1"
                End If
                If Not MovimentoBomba.AlteraPrecoVendaBomba(g_empresa, xData, 2, xPrecoG) Then
                    MsgBox "Erro Bico=2"
                End If
                If Not MovimentoBomba.AlteraPrecoVendaBomba(g_empresa, xData, 3, xPrecoG) Then
                    MsgBox "Erro Bico=3"
                End If
                If Not MovimentoBomba.AlteraPrecoVendaBomba(g_empresa, xData, 4, xPrecoGA) Then
                    MsgBox "Erro Bico=4"
                End If
                If Not MovimentoBomba.AlteraPrecoVendaBomba(g_empresa, xData, 5, xPrecoG) Then
                    MsgBox "Erro Bico=5"
                End If
                If Not MovimentoBomba.AlteraPrecoVendaBomba(g_empresa, xData, 6, xPrecoA) Then
                    MsgBox "Erro Bico=6"
                End If
                If Not MovimentoBomba.AlteraPrecoVendaBomba(g_empresa, xData, 7, xPrecoD) Then
                    MsgBox "Erro Bico=7"
                End If
                If Not MovimentoBomba.AlteraPrecoVendaBomba(g_empresa, xData, 8, xPrecoG) Then
                    MsgBox "Erro Bico=8"
                End If
                xData = rsMovBomba("Data").Value
            End If
            If rsMovBomba("Tipo de Combustivel").Value = "A " Then
                xPrecoA = rsMovBomba("Preco de Venda").Value
            ElseIf rsMovBomba("Tipo de Combustivel").Value = "DA" Then
                xPrecoD = rsMovBomba("Preco de Venda").Value
            ElseIf rsMovBomba("Tipo de Combustivel").Value = "G " Then
                xPrecoG = rsMovBomba("Preco de Venda").Value
            ElseIf rsMovBomba("Tipo de Combustivel").Value = "GA" Then
                xPrecoGA = rsMovBomba("Preco de Venda").Value
            Else
                MsgBox "ERRO "
            End If
            rsMovBomba.MoveNext
        Loop
        If Not MovimentoBomba.AlteraPrecoVendaBomba(g_empresa, xData, 1, xPrecoGA) Then
            MsgBox "Erro Bico=1"
        End If
        If Not MovimentoBomba.AlteraPrecoVendaBomba(g_empresa, xData, 2, xPrecoG) Then
            MsgBox "Erro Bico=2"
        End If
        If Not MovimentoBomba.AlteraPrecoVendaBomba(g_empresa, xData, 3, xPrecoG) Then
            MsgBox "Erro Bico=3"
        End If
        If Not MovimentoBomba.AlteraPrecoVendaBomba(g_empresa, xData, 4, xPrecoGA) Then
            MsgBox "Erro Bico=4"
        End If
        If Not MovimentoBomba.AlteraPrecoVendaBomba(g_empresa, xData, 5, xPrecoG) Then
            MsgBox "Erro Bico=5"
        End If
        If Not MovimentoBomba.AlteraPrecoVendaBomba(g_empresa, xData, 6, xPrecoA) Then
            MsgBox "Erro Bico=6"
        End If
        If Not MovimentoBomba.AlteraPrecoVendaBomba(g_empresa, xData, 7, xPrecoD) Then
            MsgBox "Erro Bico=7"
        End If
        If Not MovimentoBomba.AlteraPrecoVendaBomba(g_empresa, xData, 8, xPrecoG) Then
            MsgBox "Erro Bico=8"
        End If
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
End Sub
Private Sub zzTransfereVendaBicos()
    Dim rsMovBomba As New adodb.Recordset
    Dim xSQL As String
    Dim xBicoOrigem As Integer
    Dim xBicoDestino As Integer
    
    Exit Sub
    xBicoOrigem = 2
    xBicoDestino = 3
    xSQL = "SELECT *"
    xSQL = xSQL & "  FROM Movimento_Bomba_LMC"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & "   AND Data >= " & preparaData(CDate("01/11/2002"))
    xSQL = xSQL & "   AND [Codigo da Bomba] = " & xBicoOrigem
    xSQL = xSQL & " ORDER BY Data"
    Set rsMovBomba = Conectar.RsConexao(xSQL)
    If rsMovBomba.RecordCount > 0 Then
        Do Until rsMovBomba.EOF
            If MovimentoBomba.LocalizarCodigo(g_empresa, rsMovBomba("Data").Value, rsMovBomba("Periodo").Value, xBicoDestino, rsMovBomba("SubCaixa").Value) Then
                MovimentoBomba.QuantidadeSaida = MovimentoBomba.QuantidadeSaida + rsMovBomba("Quantidade da Saida").Value
                MovimentoBomba.Encerrante = MovimentoBomba.Encerrante + rsMovBomba("Quantidade da Saida").Value
                If Not MovimentoBomba.Alterar(g_empresa, rsMovBomba("Data").Value, rsMovBomba("Periodo").Value, xBicoDestino, rsMovBomba("SubCaixa").Value) Then
                    MsgBox "Erro Bico=2"
                End If
            Else
                MsgBox "Erro ao localizar movimento de bomba.", vbCritical, "Erro de Integridade!"
            End If
            rsMovBomba.MoveNext
        Loop
    End If
    rsMovBomba.Close
    Set rsMovBomba = Nothing
End Sub

