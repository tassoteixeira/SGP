VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form movimento_composicao_caixa 
   Caption         =   "Movimento da Composicao do Caixa"
   ClientHeight    =   7215
   ClientLeft      =   1410
   ClientTop       =   1545
   ClientWidth     =   8955
   Icon            =   "movimento_composicao_caixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "movimento_composicao_caixa.frx":030A
   ScaleHeight     =   7215
   ScaleWidth      =   8955
   Begin VB.Frame frmDados 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8715
      Begin VB.TextBox txt_celula 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2100
         TabIndex        =   14
         Top             =   3900
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txt_funcionario 
         Height          =   285
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   11
         Top             =   900
         Width           =   795
      End
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   6420
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   180
         Width           =   495
      End
      Begin VB.ComboBox cbo_tipo_movimento 
         Height          =   315
         Left            =   6420
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   540
         Width           =   2175
      End
      Begin VB.TextBox txt_numero_ilha 
         Height          =   300
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   7
         Top             =   540
         Width           =   255
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   3180
         Picture         =   "movimento_composicao_caixa.frx":0750
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   180
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   300
         Left            =   2040
         TabIndex        =   2
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc adodcFuncionario 
         Height          =   330
         Left            =   4320
         Top             =   960
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
         Bindings        =   "movimento_composicao_caixa.frx":1A2A
         Height          =   315
         Left            =   2880
         TabIndex        =   12
         Top             =   900
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboFuncionario"
      End
      Begin MSAdodcLib.Adodc adodc_composicao_caixa 
         Height          =   330
         Left            =   2220
         Top             =   5640
         Visible         =   0   'False
         Width           =   3195
         _ExtentX        =   5636
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
         Caption         =   "adodc_composicao_caixa"
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
      Begin MSDataListLib.DataCombo dtcbo_composicao_celula 
         Bindings        =   "movimento_composicao_caixa.frx":1A49
         Height          =   315
         Left            =   2820
         TabIndex        =   15
         Top             =   4380
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_produto_celula"
      End
      Begin MSFlexGridLib.MSFlexGrid fgd_composicao_caixa 
         Height          =   4275
         Left            =   0
         TabIndex        =   13
         Top             =   1320
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   7541
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   12632256
         AllowUserResizing=   1
      End
      Begin VB.Label Label6 
         Caption         =   "Total"
         Height          =   315
         Left            =   6540
         TabIndex        =   16
         Top             =   5700
         Width           =   735
      End
      Begin VB.Label lbl_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7320
         TabIndex        =   17
         Top             =   5700
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "&Data do movimento"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Período"
         Height          =   315
         Index           =   6
         Left            =   4500
         TabIndex        =   4
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Funcionário"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   10
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Tipo do movimento"
         Height          =   315
         Index           =   7
         Left            =   4500
         TabIndex        =   8
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "Número da &Ilha"
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   540
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "movimento_composicao_caixa.frx":1A6E
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Cria um novo registro."
      Top             =   6300
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "movimento_composicao_caixa.frx":3100
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Altera o registro atual."
      Top             =   6300
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "movimento_composicao_caixa.frx":45FA
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Exclui o registro atual."
      Top             =   6300
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "movimento_composicao_caixa.frx":5C8C
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   6300
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "movimento_composicao_caixa.frx":70FE
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   6300
      Width           =   795
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   6660
      TabIndex        =   25
      Top             =   6180
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "movimento_composicao_caixa.frx":8790
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "movimento_composicao_caixa.frx":9C8A
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "movimento_composicao_caixa.frx":B184
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "movimento_composicao_caixa.frx":C5F6
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   7140
      Picture         =   "movimento_composicao_caixa.frx":DB78
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Confirma o registro atual."
      Top             =   6300
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   8040
      Picture         =   "movimento_composicao_caixa.frx":F182
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Cancela o registro atual."
      Top             =   6300
      Width           =   795
   End
End
Attribute VB_Name = "movimento_composicao_caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagMovimento As Integer
Dim lOpcao As String
Dim lQtdPeriodo As Integer

Dim lEmpresa As Integer
Dim lData As Date
Dim lIlha As Integer
Dim lPeriodo As Integer
Dim lTipoMovimento As Integer
Dim lCodigoComposicao As Integer
Dim lNumeroMovimentoCaixa(0 To 20) As Long

Const NovaLinha As String = ">*"      ' Indica uma nova linha
Private ControlVisible As Boolean     ' Se o controle esta visivel ou nao
Private LastRow As Long               ' Ultima linha em que se editou
Private LastCol As Long               ' ultima coluna em que se editou
Dim lMarcaCelula As Boolean

Private IntegracaoCaixa As New cIntegracaoCaixa
Private rsMovComposicaoCaixa As New adodb.Recordset
Private ComposicaoCaixa As New cComposicaoCaixa
Private Configuracao As New cConfiguracao
Private Funcionario As New cFuncionario
Private MovAfericao As New cMovimentoAfericao
Private MovBomba As New cMovimentoBomba
Private MovCaixa As New cMovimentoCaixa
Private MovCartaoCredito As New cMovimentoCartaoCredito
Private MovCheque As New cMovimentoCheque
Private MovChequeAvista As New cMovimentoChequeAvista
Private MovCupomFiscalItem As New cMovimentoCupomFiscalItem
Private MovDespesaCaixa As New cMovimentoDespesaCaixa
Private MovFaltaCaixa As New cMovimentoFaltaCaixa
Private MovNotaAbastecimento As New cMovimentoNotaAbastecimento
Private MovimentoComposicaoCaixa As New cMovimentoComposicaoCaixa
Private MovimentoCartaFrete As New cMovimentoCartaFrete
Private MovimentoLubrificante As New cMovimentoLubrificante
Private rsTotalizador As New adodb.Recordset
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
Function BuscaChequeAvista(x_data As Date, x_periodo As String, x_tipo_movimento As String) As Currency
    BuscaChequeAvista = MovChequeAvista.TotalPeriodo(g_empresa, x_data, x_periodo, x_tipo_movimento)
    BuscaChequeAvista = BuscaChequeAvista + MovCheque.TotalEmissaoPeriodo(g_empresa, x_data, x_data, x_periodo, x_periodo, x_tipo_movimento, "V")
End Function
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Function LoopIncluiMovimentoCaixa() As Boolean
    Dim i As Integer
    Dim xNome As String
    Dim xComplemento As String
    Dim xValor(0 To 5) As Currency
    Dim xNomeProduto(0 To 5) As String
    LoopIncluiMovimentoCaixa = False
    
    For i = 1 To (fgd_composicao_caixa.Rows - 2)
        If fgd_composicao_caixa.TextMatrix(i, 0) <> "" And fValidaValor2(fgd_composicao_caixa.TextMatrix(i, 2)) > 0 Then
            xNome = ""
            If UCase(fgd_composicao_caixa.TextMatrix(i, 1)) Like "*DINHEIRO*" Then
                xNome = "DINHEIRO"
            End If
            If UCase(fgd_composicao_caixa.TextMatrix(i, 1)) Like "*MOEDA*" Then
                xNome = "MOEDA"
            End If
            If UCase(fgd_composicao_caixa.TextMatrix(i, 1)) Like "*DESPESA*" Then
                xNome = "DESPESA CAIXA"
            End If
            If UCase(fgd_composicao_caixa.TextMatrix(i, 1)) Like "*CHEQUE*" And UCase(fgd_composicao_caixa.TextMatrix(i, 1)) Like "*VISTA*" Then
                xNome = "CHEQUE A VISTA"
            End If
            If xNome <> "" Then
                If IntegracaoCaixa.LocalizarNome(g_empresa, "COMPOSICAO CAIXA-" & xNome) Then
                    xComplemento = xNome & " " & Mid(cbo_tipo_movimento.Text, 3, Len(cbo_tipo_movimento.Text) - 2) & " - Per." & Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
                    If IncluiMovimentoCaixa(CCur(fgd_composicao_caixa.TextMatrix(i, 2)), IntegracaoCaixa.ContaDebito, IntegracaoCaixa.ContaCredito, IntegracaoCaixa.HistoricoPadrao, xComplemento, True) Then
                        LoopIncluiMovimentoCaixa = True
                        fgd_composicao_caixa.TextMatrix(i, 3) = MovCaixa.NumeroMovimento
                    Else
                        MsgBox "Não foi integrado no caixa o valor=" & CCur(fgd_composicao_caixa.TextMatrix(i, 2)), vbInformation, "Erro de Integridade"
                    End If
                Else
                    MsgBox "Não existe a integração=" & "COMPOSICAO CAIXA-" & xNome & ".", vbInformation, "Registro Inexistente"
                End If
            End If
        End If
    Next
    
    xNome = ""
    For i = 0 To 5
        xValor(i) = 0
        xNomeProduto(i) = ""
    Next
    If Val(cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)) = 1 Then
        xNome = "COMBUSTIVEIS"
        xNomeProduto(0) = "ALCOOL"
        xValor(0) = MovBomba.ValorVendaPeriodo(g_empresa, CDate(msk_data.Text), CDate(msk_data.Text), "A ", Val(cbo_periodo.Text), Val(cbo_periodo.Text))
        xNomeProduto(1) = "ALCOOL AD."
        xValor(1) = MovBomba.ValorVendaPeriodo(g_empresa, CDate(msk_data.Text), CDate(msk_data.Text), "AA", Val(cbo_periodo.Text), Val(cbo_periodo.Text))
        xNomeProduto(2) = "DIESEL"
        xValor(2) = MovBomba.ValorVendaPeriodo(g_empresa, CDate(msk_data.Text), CDate(msk_data.Text), "D ", Val(cbo_periodo.Text), Val(cbo_periodo.Text))
        xNomeProduto(3) = "DIESEL AD."
        xValor(3) = MovBomba.ValorVendaPeriodo(g_empresa, CDate(msk_data.Text), CDate(msk_data.Text), "DA", Val(cbo_periodo.Text), Val(cbo_periodo.Text))
        xNomeProduto(4) = "GASOLINA"
        xValor(4) = MovBomba.ValorVendaPeriodo(g_empresa, CDate(msk_data.Text), CDate(msk_data.Text), "G ", Val(cbo_periodo.Text), Val(cbo_periodo.Text))
        xNomeProduto(5) = "GASOLINA AD."
        xValor(5) = MovBomba.ValorVendaPeriodo(g_empresa, CDate(msk_data.Text), CDate(msk_data.Text), "GA", Val(cbo_periodo.Text), Val(cbo_periodo.Text))
    ElseIf Val(cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)) = 2 Then
        xNome = "LUBRIFICANTES"
        xValor(0) = MovimentoLubrificante.TotalPeriodo(g_empresa, CDate(msk_data.Text), cbo_periodo.Text, 3)
    End If
    For i = 0 To 5
        If xNome <> "" And xValor(i) > 0 Then
            If IntegracaoCaixa.LocalizarNome(g_empresa, "VENDA DE " & xNome) Then
                xComplemento = xNome & " " & Mid(cbo_tipo_movimento.Text, 3, Len(cbo_tipo_movimento.Text) - 2) & " - Per." & Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
                If xNome = "COMBUSTIVEIS" Then
                    xComplemento = xNomeProduto(i) & " " & Mid(cbo_tipo_movimento.Text, 3, Len(cbo_tipo_movimento.Text) - 2) & " - Per." & Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
                End If
                If IncluiMovimentoCaixa(xValor(i), IntegracaoCaixa.ContaDebito, IntegracaoCaixa.ContaCredito, IntegracaoCaixa.HistoricoPadrao, xComplemento, True) Then
                    LoopIncluiMovimentoCaixa = True
                    'fgd_composicao_caixa.TextMatrix(i, 3) = MovCaixa.NumeroMovimento
                Else
                    MsgBox "Não foi integrado no caixa o valor=" & CCur(fgd_composicao_caixa.TextMatrix(i, 2)), vbInformation, "Erro de Integridade"
                End If
            Else
                MsgBox "Não existe a integração=" & "VENDA DE " & xNome & ".", vbInformation, "Registro Inexistente"
            End If
        End If
    Next
End Function
Function IncluiMovimentoCaixa(ByVal pValor As Currency, ByVal pNumeroContaDebito As String, ByVal pNumeroContaCredito As String, ByVal pCodigoHistorico As Integer, ByVal pComplemento As String, ByVal pFluxoCaixa As Boolean) As Boolean
    IncluiMovimentoCaixa = False
    MovCaixa.Empresa = g_empresa
    MovCaixa.Data = CDate(msk_data.Text)
    MovCaixa.NumeroMovimento = 1
    MovCaixa.valor = pValor
    MovCaixa.NumeroDocumento = ""
    MovCaixa.CodigoHistorico = pCodigoHistorico
    MovCaixa.Complemento = pComplemento
    MovCaixa.NumeroContaDebito = pNumeroContaDebito
    MovCaixa.NumeroContaCredito = pNumeroContaCredito
    MovCaixa.TipoMovimento = 2
    MovCaixa.FluxoCaixa = pFluxoCaixa
    MovCaixa.CodigoUsuario = g_usuario
    If MovCaixa.Incluir > 0 Then
        IncluiMovimentoCaixa = True
    End If
End Function
Private Sub AtribuiValorCelula()
    Dim Texto As String
    '
    txt_celula.Visible = False
    ControlVisible = False
    '
    ' atribuir o texto anterior a celula
    Select Case LastCol
      Case 4 To 7
        'notas menores que 5 muda cor fonte para vermelho, demais azul
        Texto = txt_celula.Text
        fgd_composicao_caixa.TextMatrix(LastRow, LastCol) = Texto
        'If Val(fgd_composicao_caixa.Text) < 6 Then
        '     fgd_composicao_caixa.CellForeColor = vbRed
        'Else
        '     fgd_composicao_caixa.CellForeColor = vbBlue
        'End If
      Case Else
        'If LastRow = 0 And LastCol = 0 Then
            LastRow = fgd_composicao_caixa.Row
            LastCol = fgd_composicao_caixa.Col
        'End If
      
        Texto = txt_celula.Text
        fgd_composicao_caixa.TextMatrix(LastRow, LastCol) = Texto
    End Select
End Sub
Private Sub AtualizaConstantes()
    If Configuracao.LocalizarCodigo(g_empresa) Then
        lQtdPeriodo = Configuracao.QuantidadePeriodos
    Else
        lQtdPeriodo = 1
    End If
End Sub
Private Sub AtualizaGrid()
    Dim xTotal As Currency
    Dim i As Integer
    Dim xSQL As String
    
    LimpaGrid
    i = 0
    fgd_composicao_caixa.Visible = False
    xTotal = 0
    
    xSQL = ""
    xSQL = xSQL & "   SELECT [Codigo da Composicao], Valor, Ordem, [Numero do Movimento do Caixa]"
    xSQL = xSQL & "     FROM Movimento_Composicao_Caixa, Composicao_Caixa"
    xSQL = xSQL & "    WHERE Empresa = " & g_empresa
    xSQL = xSQL & "      AND Data = " & preparaData(CDate(msk_data.Text))
    xSQL = xSQL & "      AND Periodo = " & Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
    xSQL = xSQL & "      AND [Numero da Ilha] = " & Val(txt_numero_ilha)
    xSQL = xSQL & "      AND [Tipo do Movimento] = " & Val(cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex))
    xSQL = xSQL & "      AND Composicao_Caixa.Codigo = [Codigo da Composicao]"
    xSQL = xSQL & " ORDER BY Composicao_Caixa.Ordem"
    Set rsMovComposicaoCaixa = New adodb.Recordset
    Set rsMovComposicaoCaixa = Conectar.RsConexao(xSQL)
    If Not rsMovComposicaoCaixa.EOF Then
        Do Until rsMovComposicaoCaixa.EOF
            i = i + 1
            fgd_composicao_caixa.Rows = fgd_composicao_caixa.Rows + 1
            fgd_composicao_caixa.Row = i
            fgd_composicao_caixa.Col = 0
            fgd_composicao_caixa.Text = Format(rsMovComposicaoCaixa("Codigo da Composicao").Value, "#,##0")
            fgd_composicao_caixa.Col = 1
            If ComposicaoCaixa.LocalizarCodigo(Val(rsMovComposicaoCaixa("Codigo da Composicao").Value)) Then
                fgd_composicao_caixa.Text = ComposicaoCaixa.Nome
            Else
                fgd_composicao_caixa.Text = "** Não Cadastrada **"
            End If
            fgd_composicao_caixa.Col = 2
            fgd_composicao_caixa.Text = Format(rsMovComposicaoCaixa("Valor").Value, "###,###,##0.00")
            fgd_composicao_caixa.Col = 3
            fgd_composicao_caixa.Text = rsMovComposicaoCaixa("Numero do Movimento do Caixa").Value
            xTotal = xTotal + rsMovComposicaoCaixa("Valor").Value
            rsMovComposicaoCaixa.MoveNext
        Loop
    End If
    
    rsMovComposicaoCaixa.Close
    Set rsMovComposicaoCaixa = Nothing
    fgd_composicao_caixa.Row = 1
    fgd_composicao_caixa.Col = 2
    fgd_composicao_caixa.Visible = True
    lbl_total.Caption = Format(xTotal, "###,###,##0.00")
    frmDados.Enabled = False
End Sub
Private Sub AtualTabe()
    Dim i As Integer
    For i = 1 To (fgd_composicao_caixa.Rows - 2)
        If fgd_composicao_caixa.TextMatrix(i, 0) <> "" And fValidaValor2(fgd_composicao_caixa.TextMatrix(i, 2)) > 0 Then
            MovimentoComposicaoCaixa.Empresa = g_empresa
            MovimentoComposicaoCaixa.Data = Format(msk_data.Text, "dd/mm/yyyy")
            MovimentoComposicaoCaixa.Periodo = Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
            MovimentoComposicaoCaixa.NumeroIlha = Val(txt_numero_ilha)
            MovimentoComposicaoCaixa.TipoMovimento = Val(cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex))
            MovimentoComposicaoCaixa.CodigoFuncionario = Val(dtcboFuncionario.BoundText)
            MovimentoComposicaoCaixa.CodigoComposicao = CLng(fgd_composicao_caixa.TextMatrix(i, 0))
            MovimentoComposicaoCaixa.valor = CCur(fgd_composicao_caixa.TextMatrix(i, 2))
            MovimentoComposicaoCaixa.NumeroMovimentoCaixa = Val(fgd_composicao_caixa.TextMatrix(i, 3))
            If MovimentoComposicaoCaixa.Incluir Then
                lData = msk_data.Text
                lPeriodo = cbo_periodo.ItemData(cbo_periodo.ListIndex)
                lIlha = Val(txt_numero_ilha.Text)
                lTipoMovimento = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
                lCodigoComposicao = CLng(fgd_composicao_caixa.TextMatrix(i, 0))
            Else
                MsgBox "Registro não foi gravado!", vbInformation, "Erro Interno"
            End If
        End If
    Next
End Sub
Private Sub AtualTela()
    Dim i As Integer
    lData = MovimentoComposicaoCaixa.Data
    lPeriodo = MovimentoComposicaoCaixa.Periodo
    lIlha = MovimentoComposicaoCaixa.NumeroIlha
    lTipoMovimento = MovimentoComposicaoCaixa.TipoMovimento
    lCodigoComposicao = MovimentoComposicaoCaixa.CodigoComposicao
    
    msk_data.Text = MovimentoComposicaoCaixa.Data
    cbo_periodo.ListIndex = -1
    For i = 0 To cbo_periodo.ListCount - 1
        If cbo_periodo.ItemData(i) = MovimentoComposicaoCaixa.Periodo Then
            cbo_periodo.ListIndex = i
            Exit For
        End If
    Next
    txt_numero_ilha.Text = MovimentoComposicaoCaixa.NumeroIlha
    cbo_tipo_movimento.ListIndex = -1
    For i = 0 To cbo_tipo_movimento.ListCount - 1
        If cbo_tipo_movimento.ItemData(i) = MovimentoComposicaoCaixa.TipoMovimento Then
            cbo_tipo_movimento.ListIndex = i
            Exit For
        End If
    Next
    txt_funcionario.Text = MovimentoComposicaoCaixa.CodigoFuncionario
    dtcboFuncionario.BoundText = MovimentoComposicaoCaixa.CodigoFuncionario
    frmDados.Enabled = False
    VerificaLiberacaoDigitacao
End Sub
Private Sub AutomatizaGridInclusao()
    Dim i As Integer
    Dim xSQL As String
    For i = 1 To (fgd_composicao_caixa.Rows - 2)
        If fgd_composicao_caixa.TextMatrix(i, 0) <> "" Then
            Exit Sub
        End If
    Next
    i = 0
    xSQL = ""
    xSQL = xSQL & "SELECT Codigo, Nome, Configuracao, Ordem"
    xSQL = xSQL & "  FROM Composicao_Caixa"
    xSQL = xSQL & " WHERE Ativo = " & preparaBooleano(True)
    xSQL = xSQL & " ORDER BY Ordem"
    Set rsMovComposicaoCaixa = New adodb.Recordset
    Set rsMovComposicaoCaixa = Conectar.RsConexao(xSQL)
    If Not rsMovComposicaoCaixa.EOF Then
        Do Until rsMovComposicaoCaixa.EOF
            i = i + 1
            fgd_composicao_caixa.Rows = fgd_composicao_caixa.Rows + 1
            fgd_composicao_caixa.Row = i
            fgd_composicao_caixa.Col = 0
            fgd_composicao_caixa.Text = Format(rsMovComposicaoCaixa("Codigo").Value, "#,##0")
            fgd_composicao_caixa.Col = 1
            fgd_composicao_caixa.Text = rsMovComposicaoCaixa("Nome").Value
            fgd_composicao_caixa.Col = 2
            fgd_composicao_caixa.Text = Format(BuscaTotais(rsMovComposicaoCaixa("Configuracao").Value), "###,###,##0.00")
            fgd_composicao_caixa.Col = 3
            fgd_composicao_caixa.Text = "0"
            rsMovComposicaoCaixa.MoveNext
        Loop
    End If
    rsMovComposicaoCaixa.Close
    Set rsMovComposicaoCaixa = Nothing
    TotalizaGrid
    fgd_composicao_caixa.Row = 1
    fgd_composicao_caixa.Col = 2
End Sub
Function BuscaTotais(ByVal xConfiguracao As String) As Currency
    BuscaTotais = 0
    If Mid(xConfiguracao, 1, 3) = "CAR" Then
        BuscaTotais = MovCartaoCredito.TotalPeriodo(g_empresa, CDate(msk_data.Text), cbo_periodo.Text, cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex), Val(Mid(xConfiguracao, 4, 2)))
    ElseIf xConfiguracao = "AFERI" Then
        If Mid(cbo_tipo_movimento.Text, 1, 1) = 1 Then
            BuscaTotais = MovAfericao.TotalPeriodo(g_empresa, CDate(msk_data.Text), Val(cbo_periodo.Text), False)
        End If
    ElseIf xConfiguracao = "CHVIS" Then
        BuscaTotais = BuscaChequeAvista(CDate(msk_data.Text), Val(cbo_periodo.Text), cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex))
    ElseIf xConfiguracao = "CHPRE" Then
        BuscaTotais = MovCheque.TotalEmissaoPeriodo(g_empresa, CDate(msk_data.Text), CDate(msk_data.Text), Val(cbo_periodo.Text), Val(cbo_periodo.Text), cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex), "P")
    ElseIf xConfiguracao = "DESPC" Then
        BuscaTotais = MovDespesaCaixa.TotalPeriodo(g_empresa, CDate(msk_data.Text), Val(cbo_periodo.Text), Val(cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)))
    ElseIf xConfiguracao = "VLTRE" Then
        BuscaTotais = BuscaValeTrocoEmitido(CDate(msk_data.Text), Val(cbo_periodo.Text), Val(cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)))
    ElseIf xConfiguracao = "VLTRR" Then
        BuscaTotais = BuscaValeTrocoRecebido(CDate(msk_data.Text), Val(cbo_periodo.Text), Val(cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)))
    ElseIf xConfiguracao = "NOTAA" Then
        If cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex) > -1 Then
            BuscaTotais = MovNotaAbastecimento.TotalPeriodo(g_empresa, CDate(msk_data.Text), cbo_periodo.Text, cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex))
        End If
    ElseIf xConfiguracao = "TRANS" Then
        If Mid(cbo_tipo_movimento.Text, 1, 1) = 1 Then
            BuscaTotais = MovAfericao.TotalPeriodo(g_empresa, CDate(msk_data.Text), Val(cbo_periodo.Text), True)
        End If
    ElseIf xConfiguracao = "DESPP" Then
        BuscaTotais = MovCupomFiscalItem.TotalDesconto(g_empresa, CDate(msk_data.Text), CDate(msk_data.Text), Val(cbo_periodo.Text), Val(cbo_periodo.Text), 0, True)
    ElseIf xConfiguracao = "CFRET" Then
        BuscaTotais = MovimentoCartaFrete.TotalCartaFrete(g_empresa, CDate(msk_data.Text), CDate(msk_data.Text), Val(cbo_periodo.Text), Val(cbo_periodo.Text), cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex))
    ElseIf xConfiguracao = "FALCX" Then
        BuscaTotais = MovFaltaCaixa.TotalFaltaCaixa(g_empresa, CDate(msk_data.Text), CDate(msk_data.Text), cbo_periodo.Text, Val(txt_numero_ilha.Text))
    ElseIf xConfiguracao = "VALEF" Then
        If Val(cbo_tipo_movimento.Text) = 1 Then
            BuscaTotais = MovFaltaCaixa.TotalValeCaixa(g_empresa, CDate(msk_data.Text), CDate(msk_data.Text), cbo_periodo.Text, Val(txt_numero_ilha.Text))
        End If
    End If
End Function
Function BuscaValeTrocoEmitido(x_data As Date, x_periodo As Integer, x_tipo_movimento As Integer) As Currency
    Dim xSQL As String
    BuscaValeTrocoEmitido = 0
    xSQL = ""
    xSQL = xSQL & "   SELECT SUM(Valor) AS Total"
    xSQL = xSQL & "     FROM Movimento_Vale_Abastecimento_Emitido"
    xSQL = xSQL & "    WHERE Empresa = " & g_empresa
    xSQL = xSQL & "      AND Data = " & preparaData(x_data)
    xSQL = xSQL & "      AND Periodo = " & x_periodo
    xSQL = xSQL & "      AND [Tipo de Movimento] = " & x_tipo_movimento
    Set rsTotalizador = New adodb.Recordset
    Set rsTotalizador = Conectar.RsConexao(xSQL)
    If Not rsTotalizador.EOF Then
        If Not IsNull(rsTotalizador("Total").Value) Then
            BuscaValeTrocoEmitido = rsTotalizador("Total").Value
        End If
    End If
    rsTotalizador.Close
    Set rsTotalizador = Nothing
End Function
Function BuscaValeTrocoRecebido(x_data As Date, x_periodo As Integer, x_tipo_movimento As Integer) As Currency
    Dim xSQL As String
    BuscaValeTrocoRecebido = 0
    xSQL = ""
    xSQL = xSQL & "   SELECT SUM(Valor) AS Total"
    xSQL = xSQL & "     FROM Movimento_Vale_Abastecimento_Recebido"
    xSQL = xSQL & "    WHERE Empresa = " & g_empresa
    xSQL = xSQL & "      AND [Data do Recebimento] = " & preparaData(x_data)
    xSQL = xSQL & "      AND [Periodo do Recebimento]= " & x_periodo
    xSQL = xSQL & "      AND [Tipo de Movimento do Recebimento] = " & x_tipo_movimento
    Set rsTotalizador = New adodb.Recordset
    Set rsTotalizador = Conectar.RsConexao(xSQL)
    If Not rsTotalizador.EOF Then
        If Not IsNull(rsTotalizador("Total").Value) Then
            BuscaValeTrocoRecebido = rsTotalizador("Total").Value
        End If
    End If
    rsTotalizador.Close
    Set rsTotalizador = Nothing
End Function
Private Sub CarregaNumeroMovimentoCaixa()
    Dim i As Integer
    For i = 0 To 20
        lNumeroMovimentoCaixa(i) = 0
    Next
    For i = 1 To (fgd_composicao_caixa.Rows - 2)
        lNumeroMovimentoCaixa(i) = Val(fgd_composicao_caixa.TextMatrix(i, 3))
        fgd_composicao_caixa.TextMatrix(i, 3) = "0"
    Next
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
    Dim i As Integer
    Dim xNome As String
    Dim xComplemento As String
    For i = 1 To 20
        If lNumeroMovimentoCaixa(i) > 0 Then
            If Not MovCaixa.Excluir(g_empresa, lData, lNumeroMovimentoCaixa(i)) Then
                MsgBox "Não foi excluído o movimento do caixa!", vbInformation, "Erro de Integridade."
            End If
        End If
    Next
    
    'Exclui Movimento "Especial"
    If MovimentoComposicaoCaixa.TipoMovimento = 1 Then
        xNome = "COMBUSTIVEIS"
    ElseIf MovimentoComposicaoCaixa.TipoMovimento = 2 Then
        xNome = "LUBRIFICANTES"
    End If
    If IntegracaoCaixa.LocalizarNome(g_empresa, "VENDA DE " & xNome) Then
        Do Until MovCaixa.LocalizarRegistroEspecial(g_empresa, CDate(msk_data.Text), IntegracaoCaixa.ContaCredito, "C") = False
            lNumeroMovimentoCaixa(0) = MovCaixa.NumeroMovimento
            If Not MovCaixa.Excluir(g_empresa, CDate(msk_data.Text), lNumeroMovimentoCaixa(0)) Then
                MsgBox "Não foi excluído o movimento do caixa!", vbInformation, "Erro de Integridade."
            End If
        Loop
    Else
        MsgBox "Não existe a integração=" & "VENDA DE " & xNome & ".", vbInformation, "Registro Inexistente"
    End If
End Sub
Private Sub ExibirCelula()
    Static OK As Boolean
    '
    ' Se for celula fixa , sair
    If fgd_composicao_caixa.Col <= fgd_composicao_caixa.FixedCols - 1 Or fgd_composicao_caixa.Row <= fgd_composicao_caixa.FixedRows - 1 Then
       Exit Sub
    End If
    
    If OK Then Exit Sub
    OK = True
    '
    txt_celula.Visible = False
    '
    LastRow = fgd_composicao_caixa.Row
    LastCol = fgd_composicao_caixa.Col
    If LastCol = 0 Then
        txt_celula.MaxLength = 4
    ElseIf LastCol = 2 Then
        txt_celula.MaxLength = 10
    End If
    
    '
    ' Nova Celula
    'With fgd_composicao_caixa
    '  If .TextMatrix(LastRow, 0) = NovaLinha Then
    '    .Rows = .Rows + 1
    '    .TextMatrix(LastRow, 0) = LastRow
    '    .TextMatrix(.Rows - 1, 0) = NovaLinha
    '  End If
    'End With
    '
    Select Case LastCol
        Case Else
        txt_celula.Move fgd_composicao_caixa.CellLeft - Screen.TwipsPerPixelX, fgd_composicao_caixa.CellTop + 1300 - Screen.TwipsPerPixelY, fgd_composicao_caixa.CellWidth + Screen.TwipsPerPixelX * 2, fgd_composicao_caixa.CellHeight + Screen.TwipsPerPixelY * 2
        txt_celula.Text = fgd_composicao_caixa.Text
        'If Len(fgd_composicao_caixa.Text) = 0 Then
        '   If LastRow > 1 Then
        '       txt_celula.Text = fgd_composicao_caixa.TextMatrix(LastRow - 1, LastCol)
        '   End If
        'End If
        txt_celula.Visible = True
        If txt_celula.Visible Then
          txt_celula.ZOrder
          txt_celula.SetFocus
        End If
    End Select
    ControlVisible = True
    OK = False
End Sub
Private Sub ExibirComposicaoCaixaCelula()
    Dim i As Integer
    Static OK As Boolean
    '
    ' Se for celula fixa , sair
    If fgd_composicao_caixa.Col <= fgd_composicao_caixa.FixedCols - 1 Or fgd_composicao_caixa.Row <= fgd_composicao_caixa.FixedRows - 1 Then
       Exit Sub
    End If
    
    If OK Then Exit Sub
    OK = True
    '
    dtcbo_composicao_celula.Visible = False
    '
    LastRow = fgd_composicao_caixa.Row
    LastCol = fgd_composicao_caixa.Col
    '
    ' Nova Celula
    'With fgd_composicao_caixa
    '  If .TextMatrix(LastRow, 0) = NovaLinha Then
    '    .Rows = .Rows + 1
    '    .TextMatrix(LastRow, 0) = LastRow
    '    .TextMatrix(.Rows - 1, 0) = NovaLinha
    '  End If
    'End With
    '
    Select Case LastCol
        Case Else
        dtcbo_composicao_celula.Left = fgd_composicao_caixa.CellLeft - Screen.TwipsPerPixelX
        dtcbo_composicao_celula.Top = fgd_composicao_caixa.CellTop + 1300 - Screen.TwipsPerPixelY
        dtcbo_composicao_celula.Width = fgd_composicao_caixa.CellWidth + Screen.TwipsPerPixelX * 2
        'dtcbo_composicao_celula.Move fgd_composicao_caixa.CellLeft - Screen.TwipsPerPixelX, fgd_composicao_caixa.CellTop + 2095 - Screen.TwipsPerPixelY, fgd_composicao_caixa.CellWidth + Screen.TwipsPerPixelX * 2, fgd_composicao_caixa.CellHeight + Screen.TwipsPerPixelY * 2
        dtcbo_composicao_celula.BoundText = ""
        If Val(fgd_composicao_caixa.TextMatrix(LastRow, 0)) > 0 Then
            dtcbo_composicao_celula.BoundText = Val(fgd_composicao_caixa.TextMatrix(LastRow, 0))
            'For i = 0 To dtcbo_composicao_celula.ListCount
            '    If Val(fgd_composicao_caixa.TextMatrix(LastRow, 0)) > dtcbo_composicao_celula.ListCount Then
            '        Exit For
            '    End If
            '    If dtcbo_composicao_celula.ItemData(i) = Val(fgd_composicao_caixa.TextMatrix(LastRow, 0)) Then
            '        dtcbo_composicao_celula.ListIndex = i
            '        Exit For
            '    End If
            'Next
        End If
        'If Len(fgd_composicao_caixa.Text) = 0 Then
        '   If LastRow > 1 Then
        '       txt_celula.Text = fgd_composicao_caixa.TextMatrix(LastRow - 1, LastCol)
        '   End If
        'End If
        dtcbo_composicao_celula.Visible = True
        If dtcbo_composicao_celula.Visible Then
          dtcbo_composicao_celula.ZOrder
          dtcbo_composicao_celula.SetFocus
        End If
    End Select
    ControlVisible = True
    OK = False
End Sub
Private Sub Finaliza()
    Set IntegracaoCaixa = Nothing
    Set MovAfericao = Nothing
    Set MovBomba = Nothing
    Set MovCaixa = Nothing
    Set MovCartaoCredito = Nothing
    Set MovCheque = Nothing
    Set MovChequeAvista = Nothing
    Set MovCupomFiscalItem = Nothing
    Set MovDespesaCaixa = Nothing
    Set MovFaltaCaixa = Nothing
    Set MovNotaAbastecimento = Nothing
    Set MovimentoCartaFrete = Nothing
    Set MovimentoLubrificante = Nothing
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
    cbo_tipo_movimento.AddItem "3 Caixa da Conveniência"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 3
End Sub
Private Sub PreencheDtCboComposicaoCaixa()
    Dim xSQL As String
    xSQL = ""
    xSQL = xSQL & "SELECT Codigo, Nome"
    xSQL = xSQL & "  FROM Composicao_Caixa"
    'xSQL = xSQL & "  WHERE Inativo = FALSE"
    'xSQL = xSQL & "  AND [Codigo do Grupo] > " & xGrupo
    xSQL = xSQL & " ORDER BY Ordem"
'    adodc_composicao_caixa.ConnectionString = gConnectionString
'    adodc_composicao_caixa.RecordSource = xSQL
'    adodc_composicao_caixa.Refresh
    Set adodc_composicao_caixa.Recordset = Conectar.RsConexao(xSQL)
End Sub
Private Sub cbo_periodo_GotFocus()
    SendMessageLong cbo_periodo.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_numero_ilha.SetFocus
    End If
End Sub
Private Sub cbo_tipo_movimento_GotFocus()
    SendMessageLong cbo_tipo_movimento.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_tipo_movimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_funcionario.SetFocus
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
    Call CarregaNumeroMovimentoCaixa
    fgd_composicao_caixa.Col = 2
    fgd_composicao_caixa.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If MovimentoComposicaoCaixa.LocalizarAnterior Then
        AtualTela
        AtualizaGrid
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If MovimentoComposicaoCaixa.LocalizarCodigo(g_empresa, lData, lIlha, lPeriodo, lTipoMovimento, lCodigoComposicao) Then
        AtualTela
        AtualizaGrid
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
    msk_data.Text = "__/__/____"
    cbo_periodo.ListIndex = -1
    cbo_tipo_movimento.ListIndex = -1
    lbl_total.Caption = ""
    txt_funcionario.Text = ""
    dtcboFuncionario.BoundText = ""
    LimpaGrid
End Sub
Private Sub LimpaGrid()
    Dim x_sql As String
    Dim i As Integer
    fgd_composicao_caixa.WordWrap = True
    fgd_composicao_caixa.Rows = 2
    fgd_composicao_caixa.Row = 1
    For i = 0 To 2
        fgd_composicao_caixa.Col = i
        fgd_composicao_caixa.Text = ""
    Next
    fgd_composicao_caixa.RowHeight(0) = 500
    fgd_composicao_caixa.Row = 0
    i = 0
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Código"
    fgd_composicao_caixa.ColWidth(i) = 900
    fgd_composicao_caixa.ColAlignment(i) = 4
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Nome da Composicao"
    fgd_composicao_caixa.ColWidth(i) = 5610
    fgd_composicao_caixa.ColAlignment(i) = 1
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Valor"
    fgd_composicao_caixa.ColWidth(i) = 900
    fgd_composicao_caixa.ColAlignment(i) = 7
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Num.Mov.Caixa"
    fgd_composicao_caixa.ColWidth(i) = 900
    fgd_composicao_caixa.ColAlignment(i) = 7
    
    'fgd_composicao_caixa.ColIsVisible = (False)
    'fgd_composicao_caixa.Visible = True '.ColIsVisible(3) = False
    'fgd_composicao_caixa.ColIsVisible(3) = False
    'x'lbl_total_nota.Caption = ""
    txt_celula.Visible = False
    dtcbo_composicao_celula.Visible = False
    fgd_composicao_caixa.Row = 1
    fgd_composicao_caixa.Col = 0
    fgd_composicao_caixa.Text = ""
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data.Text = RetiraGString(1)
    cbo_periodo.SetFocus
    g_string = ""
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_funcionario) > 0 Then
        If (MsgBox("Deseja excluir estes registros?", 4 + 32 + 256, "Exclusão de Registros!")) = 6 Then
            If MovimentoComposicaoCaixa.ExcluirRegistros(g_empresa, CDate(msk_data.Text), Val(txt_numero_ilha), Val(cbo_periodo.ItemData(cbo_periodo.ListIndex)), Val(cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex))) Then
                Call CarregaNumeroMovimentoCaixa
                Call ExcluiMovimentoCaixa
                LimpaTela
                If MovimentoComposicaoCaixa.LocalizarUltimo(g_empresa) Then
                    AtualTela
                    AtualizaGrid
                Else
                    DesativaBotoes
                    cmd_novo.Enabled = True
                    cmd_sair.Enabled = True
                    cmd_novo.SetFocus
                End If
            Else
                MsgBox "Registros não excluidos!", vbInformation, "Erro de Integridade!"
            End If
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
    LimpaTela
    Inclui
    frmDados.Enabled = True
    If BuscaProximoCaixa Then
        cbo_periodo.SetFocus
    Else
        msk_data.SetFocus
    End If
End Sub
Private Sub cmd_novo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 16 Then
        MsgBox "PROCESSAMENTO"
        Call ProcessaComposicaoCaixa
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        If VerificaLiberacaoDigitacao2 Then
            AtivaBotoes
            If lOpcao = 1 Then
                If Not LoopIncluiMovimentoCaixa Then
                    MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
                End If
                AtualTabe
            ElseIf lOpcao = 2 Then
                Call ExcluiMovimentoCaixa
                If Not LoopIncluiMovimentoCaixa Then
                    MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
                End If
                Call MovimentoComposicaoCaixa.ExcluirRegistros(g_empresa, lData, lIlha, lPeriodo, lTipoMovimento)
                AtualTabe
            End If
            lOpcao = 0
            Call MovimentoComposicaoCaixa.LocalizarCodigo(g_empresa, lData, lIlha, lPeriodo, lTipoMovimento, lCodigoComposicao)
            AtualizaGrid
            cmd_novo.SetFocus
        End If
    End If
    Exit Sub
FileError:
    MsgBox Error
    Exit Sub
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data do movimento.", 64, "Atenção!"
        msk_data.SetFocus
    ElseIf Not ValidaPeriodo Then
        cbo_periodo.SetFocus
    ElseIf cbo_tipo_movimento.ListIndex = -1 Then
        MsgBox "Escolha o tipo de movimento.", 64, "Atenção!"
        cbo_tipo_movimento.SetFocus
    ElseIf fValidaValor2(lbl_total.Caption) = 0 Then
        MsgBox "Informe o valor de alguma composição de caixa.", 64, "Atenção!"
        fgd_composicao_caixa.SetFocus
    ElseIf Val(dtcboFuncionario.BoundText) = 0 Then
        MsgBox "Selecione um funcionario.", 64, "Atenção!"
        dtcboFuncionario.SetFocus
    ElseIf Not Val(txt_numero_ilha.Text) > 0 Then
        MsgBox "O número da ilha deve ser maior que 0.", 64, "Atenção!"
        txt_numero_ilha.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Function ValidaPeriodo() As Boolean
    ValidaPeriodo = True
    If cbo_periodo.ListIndex = -1 Then
        MsgBox "Escolha o período.", 64, "Atenção!"
        ValidaPeriodo = False
    ElseIf cbo_periodo.ListIndex = 3 And cbo_tipo_movimento.ListIndex <> 1 Then
        If (MsgBox("Tipo de movimento não aceito para este período." & Chr(10) & Chr(10) & "Deseja cadastrar mesmo assim?", vbYesNo + vbDefaultButton2, "Atenção!")) = 7 Then
            ValidaPeriodo = False
        End If
    End If
End Function
Private Sub VerificaLiberacaoDigitacao()
    Dim x_flag As Boolean
    x_flag = True
    If g_nivel_acesso > 4 Then
        If MovimentoComposicaoCaixa.Empresa < g_cfg_empresa_i Or MovimentoComposicaoCaixa.Empresa > g_cfg_empresa_f Then
            x_flag = False
        ElseIf MovimentoComposicaoCaixa.Data < g_cfg_data_i Or MovimentoComposicaoCaixa.Data > g_cfg_data_f Then
            x_flag = False
        ElseIf MovimentoComposicaoCaixa.Periodo < g_cfg_periodo_i Or MovimentoComposicaoCaixa.Periodo > g_cfg_periodo_f Then
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
    If msk_data < g_cfg_data_i Or msk_data > g_cfg_data_f Then
        MsgBox "A data do movimento deve estar entre " & Format(g_cfg_data_i, "dd/mm/yyyy") & " a " & Format(g_cfg_data_f, "dd/mm/yyyy") & ".", 64, "Digitação Não Autorizada!"
        msk_data.SetFocus
    ElseIf cbo_periodo < g_cfg_periodo_i Or cbo_periodo > g_cfg_periodo_f Then
        MsgBox "O período deve estar entre " & g_cfg_periodo_i & " ao " & g_cfg_periodo_f & ".", 64, "Digitação Não Autorizada!"
        cbo_periodo.SetFocus
    Else
        VerificaLiberacaoDigitacao2 = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_movimento_composicao_caixa.Show 1
    If Len(g_string) > 0 Then
        lData = RetiraGString(1)
        lPeriodo = RetiraGString(2)
        lIlha = RetiraGString(3)
        lTipoMovimento = RetiraGString(4)
        lCodigoComposicao = RetiraGString(5)
        If MovimentoComposicaoCaixa.LocalizarCodigo(g_empresa, lData, lIlha, lPeriodo, lTipoMovimento, lCodigoComposicao) Then
            AtualTela
            AtualizaGrid
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If MovimentoComposicaoCaixa.LocalizarPrimeiro Then
        AtualTela
        AtualizaGrid
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If MovimentoComposicaoCaixa.LocalizarProximo Then
        AtualTela
        AtualizaGrid
    Else
        MsgBox "Fim de Arquivo.", vbInformation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If MovimentoComposicaoCaixa.LocalizarUltimo(g_empresa) Then
        AtualTela
        AtualizaGrid
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Command1_Click()
    Dim sql As String
    Exit Sub
    sql = "Update Movimento_Historico "
    sql = sql & "Set Afericao = 0, Transferencia = 0"
    bd_sgp.Execute sql
End Sub
Private Sub ProximaCelula()
    If fgd_composicao_caixa.Col < fgd_composicao_caixa.Cols - 2 Then
        fgd_composicao_caixa.Col = LastCol + 1
    Else
        fgd_composicao_caixa.Col = 2
        If fgd_composicao_caixa.Row >= fgd_composicao_caixa.Rows - 1 Then
            fgd_composicao_caixa.Rows = fgd_composicao_caixa.Rows + 1
        End If
        fgd_composicao_caixa.Row = fgd_composicao_caixa.Row + 1
    End If
    fgd_composicao_caixa.SetFocus
End Sub
Private Sub TabelaFuncionarioRefresh()
    Dim xSQL As String
    xSQL = "SELECT Codigo, Nome FROM Funcionario WHERE Empresa = " & g_empresa & " AND Situacao = " & preparaTexto("A") & " AND [Periodo] < 5 ORDER BY [Nome]"
    Set adodcFuncionario.Recordset = Conectar.RsConexao(xSQL)
End Sub
Private Sub TotalizaGrid()
    Dim x_total As Currency
    Dim i As Integer
    x_total = 0
    With fgd_composicao_caixa
        For i = 1 To (.Rows - 1)
            If Len(.TextMatrix(i, 0)) > 0 Then
                x_total = x_total + fValidaValor(.TextMatrix(i, 2))
            End If
        Next
    End With
    lbl_total.Caption = Format(x_total, "###,###,##0.00")
End Sub
Private Sub dtcbo_composicao_celula_KeyPress(KeyAscii As Integer)
    ' ao pressionar ENTER aceitar a entrada de dados
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        fgd_composicao_caixa.SetFocus
    ' ESC, cancela a edição
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        dtcbo_composicao_celula.Visible = False
        ControlVisible = False
    End If
End Sub
Private Sub dtcbo_composicao_celula_LostFocus()
    Dim x_nome_grupo As String
    If dtcbo_composicao_celula.BoundText <> "" Then
        If Not ComposicaoCaixa.LocalizarCodigo(CLng(dtcbo_composicao_celula.BoundText)) Then
            MsgBox "Composicao de caixa não cadastrada!", vbInformation, "Validação Incorreta!"
            fgd_composicao_caixa.SetFocus
            Exit Sub
        Else
            fgd_composicao_caixa.TextMatrix(LastRow, 0) = Format(ComposicaoCaixa.Codigo, "###0")
            fgd_composicao_caixa.TextMatrix(LastRow, 1) = ComposicaoCaixa.Nome
            dtcbo_composicao_celula.Visible = False
            fgd_composicao_caixa.Col = 2
            'ExibirComposicaoCaixaCelula
            fgd_composicao_caixa.SetFocus
            LastCol = 1
        End If
    Else
        dtcbo_composicao_celula.Visible = False
        cmd_ok.SetFocus
        Exit Sub
    End If
    If LastCol <> 1 Then
        ProximaCelula
    End If
End Sub
Private Sub dtcboFuncionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        fgd_composicao_caixa.SetFocus
    End If
End Sub
Private Sub dtcboFuncionario_LostFocus()
    If dtcboFuncionario.BoundText <> "" Then
        txt_funcionario.Text = dtcboFuncionario.BoundText
        If lOpcao = 1 Then
            AutomatizaGridInclusao
        End If
    End If
End Sub
Private Sub fgd_composicao_caixa_Click()
    ' Quando clicar uma vez
    ' atribui o valor selecionado
    lMarcaCelula = True
    If fgd_composicao_caixa.Col = 0 Or fgd_composicao_caixa.Col = 1 Or fgd_composicao_caixa.Col = 2 Then
        LastRow = fgd_composicao_caixa.Row
        LastCol = fgd_composicao_caixa.Col
        txt_celula.Visible = False
        dtcbo_composicao_celula.Visible = False
    End If
    'AtribuiValorCelula
End Sub
Private Sub fgd_composicao_caixa_DblClick()
    'editar ao clicar duas vezes
    lMarcaCelula = True
    If fgd_composicao_caixa.Col = 0 Or fgd_composicao_caixa.Col = 2 Then
        '0 - Código da Composicao do Caixa
        '2 - Valor
        LastRow = fgd_composicao_caixa.Row
        LastCol = fgd_composicao_caixa.Col
        txt_celula.Visible = False
        dtcbo_composicao_celula.Visible = False
        ExibirCelula
    ElseIf fgd_composicao_caixa.Col = 1 Then
        'Nome da Composicao do Caixa
        LastRow = fgd_composicao_caixa.Row
        LastCol = fgd_composicao_caixa.Col
        txt_celula.Visible = False
        dtcbo_composicao_celula.Visible = False
        ExibirComposicaoCaixaCelula
    End If
End Sub
Private Sub fgd_composicao_caixa_KeyPress(KeyAscii As Integer)
    lMarcaCelula = True
    Select Case KeyAscii
    ' Editar ao teclar ENTER
    Case vbKeyReturn
        KeyAscii = 0
        If fgd_composicao_caixa.Col = 0 Or fgd_composicao_caixa.Col = 2 Then
            ExibirCelula
        ElseIf fgd_composicao_caixa.Col = 1 Then
            ExibirComposicaoCaixaCelula
        End If
    ' Cancelar ao pressionar ESC
    Case vbKeyEscape
        KeyAscii = 0
        AtribuiValorCelula
    ' Editar ao pressinar qualquer tecla
    Case 32 To 255
        lMarcaCelula = False
        If fgd_composicao_caixa.Col = 0 Or fgd_composicao_caixa.Col = 2 Then
            ExibirCelula
            With txt_celula
                If .Visible Then
                    .Text = Chr$(KeyAscii)
                    .SelStart = Len(.Text) + 1
                End If
            End With
        ElseIf fgd_composicao_caixa.Col = 1 Then
            ExibirComposicaoCaixaCelula
            'With txt_celula
            '    If .Visible Then
            '        .Text = Chr$(KeyAscii)
            '        .SelStart = Len(.Text) + 1
            '    End If
            'End With
        End If
    End Select
End Sub
Private Sub fgd_composicao_caixa_Scroll()
    ' Ver se a coluna esta visivel
    ' entao ocultar os controles
    '
    If fgd_composicao_caixa.ColIsVisible(LastCol) = False Then
        txt_celula.Visible = False
        Exit Sub
    End If
    If fgd_composicao_caixa.RowIsVisible(LastRow) = False Then
        txt_celula.Visible = False
        Exit Sub
    End If
    ' ver se estava visivel antes de ocultar
    ' e posicionar na mesma celula
    If ControlVisible Then
        ExibirCelula
    End If
End Sub
Private Sub Form_Activate()
    If g_empresa <> lEmpresa Then
        lFlagMovimento = 0
    End If
    If lFlagMovimento = 0 Then
        AtualizaConstantes
        TabelaFuncionarioRefresh
        lOpcao = 0
        lEmpresa = g_empresa
        DesativaBotoes
        If MovimentoComposicaoCaixa.LocalizarUltimo(g_empresa) Then
            AtualTela
            AtualizaGrid
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        If cmd_novo.Enabled Then
            cmd_novo.SetFocus
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
    CentraForm Me
    
    If g_nome_usuario = "L.M.C." Then
        MovAfericao.NomeTabela = "Movimento_Afericao_LMC"
        MovBomba.NomeTabela = "Movimento_Bomba_LMC"
    ElseIf UCase(g_nome_usuario) = "CUPOM FISCAL" Or fEcfInstalada Then
        MovAfericao.NomeTabela = "Movimento_Afericao"
        MovBomba.NomeTabela = "Movimento_Bomba_Cupom"
    Else
        MovAfericao.NomeTabela = "Movimento_Afericao"
        MovBomba.NomeTabela = "Movimento_Bomba"
    End If
    
    PreencheCboPeriodo
    PreencheCboTipoMovimento
    lData = "01/01/1900"
    lPeriodo = "0"
    lTipoMovimento = "0"
    Call PreencheDtCboComposicaoCaixa
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_GotFocus()
    msk_data.SelStart = 0
    msk_data.SelLength = 5
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo.SetFocus
    End If
End Sub
Private Sub txt_celula_GotFocus()
    With txt_celula
        If LastCol = 0 Then
            .MaxLength = 4
        ElseIf LastCol = 2 Then
            .MaxLength = 10
            .Text = fValidaValor(.Text)
        End If
        If lMarcaCelula Then
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
    End With
End Sub
Private Sub txt_celula_KeyPress(KeyAscii As Integer)
    ' ao pressionar ENTER aceitar a entrada de dados
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        fgd_composicao_caixa.SetFocus
    ' ESC, cancela a edição
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        txt_celula.Visible = False
        ControlVisible = False
    End If
    If LastCol = 0 Then
        Call ValidaInteiro(KeyAscii)
    ElseIf LastCol = 2 Then
        If KeyAscii = 46 Then
            KeyAscii = 44
        End If
        Call ValidaValor(KeyAscii)
    End If
End Sub
Private Sub txt_celula_LostFocus()
    'Código do Produto
    If LastCol = 0 Then
        'If Not IsNumeric(txt_celula.Text) Then
        '    MsgBox "Informe o código do serviço.", vbInformation, "Validação Incorreta!"
        '    Exit Sub
        'End If
        If Val(txt_celula.Text) > 0 Then
            If Not ComposicaoCaixa.LocalizarCodigo(Val(txt_celula.Text)) Then
                MsgBox "Composição de Caixa não cadastrada!", vbInformation, "Validação Incorreta!"
                fgd_composicao_caixa.SetFocus
                Exit Sub
            Else
                AtribuiValorCelula
                fgd_composicao_caixa.TextMatrix(LastRow, 1) = ComposicaoCaixa.Nome
                fgd_composicao_caixa.Col = 2
                fgd_composicao_caixa.SetFocus
                LastCol = 1
            End If
        ElseIf txt_celula.Text = "" Then
            AtribuiValorCelula
        Else
            AtribuiValorCelula
            fgd_composicao_caixa.TextMatrix(LastRow, 0) = ""
            fgd_composicao_caixa.TextMatrix(LastRow, 1) = ""
            fgd_composicao_caixa.TextMatrix(LastRow, 2) = ""
            TotalizaGrid
            cmd_ok.SetFocus
            Exit Sub
        End If
    ElseIf LastCol = 2 Then
        If fValidaValor(txt_celula.Text) > 0 Then
            txt_celula.Text = Format(txt_celula.Text, "##,###,##0.00")
        Else
            txt_celula.Text = "0,00"
        End If
        AtribuiValorCelula
        TotalizaGrid
    End If
    If LastCol <> 1 Then
        ProximaCelula
    End If
End Sub
Private Sub txt_funcionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboFuncionario.SetFocus
    End If
End Sub
Private Sub txt_funcionario_LostFocus()
    If Val(txt_funcionario.Text) > 0 And lOpcao > 0 Then
        If Funcionario.LocalizarCodigo(g_empresa, Val(txt_funcionario.Text)) Then
            If Funcionario.Situacao = "A" Then
                dtcboFuncionario.BoundText = Val(txt_funcionario.Text)
                fgd_composicao_caixa.SetFocus
                Exit Sub
            Else
                MsgBox "O funcionário " & Trim(Funcionario.Nome) & " está inativo.", vbInformation, "Aviso de Verificação!"
                txt_funcionario.SetFocus
            End If
        Else
            MsgBox "Funcionário não cadastro.", vbInformation, "Aviso de Verificação!"
            txt_funcionario.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_numero_ilha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo_movimento.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Function BuscaProximoCaixa() As Boolean
    Dim x_periodo As String
    Dim x_tipo_movimento As String
    BuscaProximoCaixa = False
    msk_data.Text = g_data_def - 1
    cbo_periodo.ListIndex = 0
    cbo_tipo_movimento.ListIndex = 0
    If MovimentoComposicaoCaixa.LocalizarUltimo(g_empresa) Then
        msk_data.Text = MovimentoComposicaoCaixa.Data
        x_periodo = MovimentoComposicaoCaixa.Periodo - 1
        x_tipo_movimento = MovimentoComposicaoCaixa.TipoMovimento - 1
        x_tipo_movimento = x_tipo_movimento + 1
        If x_tipo_movimento = 2 Then
            x_tipo_movimento = 0
            x_periodo = x_periodo + 1
        End If
        If x_periodo > (lQtdPeriodo - 1) Then
            msk_data.Text = MovimentoComposicaoCaixa.Data + 1
            x_periodo = 0
            x_tipo_movimento = 0
        End If
        If x_periodo = 3 Then
            x_tipo_movimento = 1
        End If
        If x_tipo_movimento = 3 Then
            x_tipo_movimento = 0
        End If
        cbo_periodo.ListIndex = x_periodo
        cbo_tipo_movimento.ListIndex = x_tipo_movimento
        BuscaProximoCaixa = True
    End If
End Function
Private Sub ProcessaComposicaoCaixa()
    Dim xData As Date
    On Error GoTo FileError
    
    xData = CDate("01/10/2004")
    If MovimentoComposicaoCaixa.LocalizarPrimeiroData(g_empresa, xData) Then
        If MovimentoComposicaoCaixa.Data >= xData Then
            AtualTela
            AtualizaGrid
            If Not LoopIncluiMovimentoCaixa Then
                MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
            Else
                Call MovimentoComposicaoCaixa.ExcluirRegistros(g_empresa, lData, lIlha, lPeriodo, lTipoMovimento)
                AtualTabe
            End If
        End If
    
        Do Until MovimentoComposicaoCaixa.LocalizarProximo = False
            If MovimentoComposicaoCaixa.Data >= xData Then
                AtualTela
                AtualizaGrid
                If Not LoopIncluiMovimentoCaixa Then
                    MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
                Else
                    Call MovimentoComposicaoCaixa.ExcluirRegistros(g_empresa, lData, lIlha, lPeriodo, lTipoMovimento)
                    AtualTabe
                End If
            End If
        Loop
    
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
    MsgBox "Processamento Concluído!"
    Exit Sub
FileError:
    MsgBox "Erro ao processar Composição de Caixa", vbInformation, "ProcessaComposicaoCaixa"
End Sub

