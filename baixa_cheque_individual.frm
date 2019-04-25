VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form baixa_cheque_individual 
   Caption         =   "Baixa de Cheque (Individual)"
   ClientHeight    =   8220
   ClientLeft      =   3615
   ClientTop       =   585
   ClientWidth     =   9195
   Icon            =   "baixa_cheque_individual.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "baixa_cheque_individual.frx":030A
   ScaleHeight     =   8220
   ScaleWidth      =   9195
   Begin VB.TextBox txtPeriodoBaixa 
      Height          =   285
      Left            =   6480
      MaxLength       =   1
      TabIndex        =   26
      Top             =   6720
      Width           =   255
   End
   Begin VB.OptionButton optEmissao 
      Caption         =   "Data de Emissão"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.OptionButton optVencimento 
      Caption         =   "Data de Vencimento"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   420
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdDataFinal 
      Height          =   315
      Left            =   8580
      Picture         =   "baixa_cheque_individual.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Selecione a data pelo calendário."
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdDataInicial 
      Height          =   315
      Left            =   5100
      Picture         =   "baixa_cheque_individual.frx":1A2A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Selecione a data pelo calendário."
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmd_extornar 
      Caption         =   "&Extornar"
      Height          =   855
      Left            =   1020
      Picture         =   "baixa_cheque_individual.frx":2D04
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Extorna o registro atual."
      Top             =   7320
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2820
      Picture         =   "baixa_cheque_individual.frx":3FDE
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   7320
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   1920
      Picture         =   "baixa_cheque_individual.frx":5670
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   7320
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   120
      Picture         =   "baixa_cheque_individual.frx":6AE2
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Altera o registro atual."
      Top             =   7320
      Width           =   795
   End
   Begin VB.Frame frmDados 
      Enabled         =   0   'False
      Height          =   3375
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   8955
      Begin MSMask.MaskEdBox mskDataBaixa 
         Height          =   300
         Left            =   2040
         TabIndex        =   24
         Top             =   3000
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
         Caption         =   "Data da Baixa"
         Height          =   315
         Index           =   18
         Left            =   120
         TabIndex        =   23
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Período da Baixa"
         Height          =   315
         Index           =   17
         Left            =   4980
         TabIndex        =   25
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Data da Custódia"
         Height          =   315
         Index           =   16
         Left            =   4980
         TabIndex        =   55
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label lblDataCustodia 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6360
         TabIndex        =   54
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Ilha"
         Height          =   315
         Index           =   15
         Left            =   7800
         TabIndex        =   53
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblIlha 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   8460
         TabIndex        =   52
         Top             =   120
         Width           =   315
      End
      Begin VB.Label Label3 
         Caption         =   "Número do Telefone"
         Height          =   315
         Index           =   14
         Left            =   120
         TabIndex        =   51
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label lblTelefone 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   50
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Número Cheque"
         Height          =   315
         Index           =   13
         Left            =   4980
         TabIndex        =   49
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Número da Conta"
         Height          =   315
         Index           =   12
         Left            =   120
         TabIndex        =   48
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblNumeroCheque 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6360
         TabIndex        =   47
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label lblConta 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   46
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Código Agencia"
         Height          =   315
         Index           =   3
         Left            =   4980
         TabIndex        =   45
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Código do Banco"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblAgencia 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6360
         TabIndex        =   43
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblBanco 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   42
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Nome do Emitente"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblEmitente 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   11
         Top             =   1920
         Width           =   5415
      End
      Begin VB.Label lblValor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   28
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblDataVencimento 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6360
         TabIndex        =   30
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblCpfCnpj 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6360
         TabIndex        =   19
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblCodigoFuncionario 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   21
         Top             =   2640
         Width           =   795
      End
      Begin VB.Label lblNomeFuncionario 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2940
         TabIndex        =   22
         Top             =   2640
         Width           =   4515
      End
      Begin VB.Label lblTipoMovimento 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   17
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblPeriodo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6360
         TabIndex        =   15
         Top             =   120
         Width           =   315
      End
      Begin VB.Label lblDataEmissao 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2040
         TabIndex        =   13
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Valor do Cheque"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Funcionário (Vendedor)"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "CPF / CNPJ"
         Height          =   315
         Index           =   10
         Left            =   4980
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo do movimento"
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Período"
         Height          =   315
         Index           =   6
         Left            =   4980
         TabIndex        =   14
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Data Vencimento"
         Height          =   315
         Index           =   5
         Left            =   4980
         TabIndex        =   29
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Data de emissão"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   6900
      TabIndex        =   37
      Top             =   7200
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "baixa_cheque_individual.frx":7FDC
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "baixa_cheque_individual.frx":955E
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "baixa_cheque_individual.frx":A9D0
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "baixa_cheque_individual.frx":BECA
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   8280
      Picture         =   "baixa_cheque_individual.frx":D3C4
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Cancela o registro atual."
      Top             =   7320
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   7380
      Picture         =   "baixa_cheque_individual.frx":E8BE
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Confirma o registro atual."
      Top             =   7320
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   2955
      Left            =   120
      TabIndex        =   10
      Top             =   780
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   5212
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
   End
   Begin MSMask.MaskEdBox mskDataInicial 
      Height          =   300
      Left            =   3960
      TabIndex        =   3
      Top             =   240
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
   Begin MSMask.MaskEdBox mskDataFinal 
      Height          =   300
      Left            =   7440
      TabIndex        =   6
      Top             =   240
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
      Caption         =   "Data &Final"
      Height          =   315
      Index           =   11
      Left            =   6300
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Data &Inicial"
      Height          =   315
      Index           =   8
      Left            =   2820
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "baixa_cheque_individual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** PENDENCIA ***
'Criar pesquisa de cheque baixado
'Integrar na movimentação do caixa
'Criar campo de Data da Baixa
'Criar campo [Numero do Movimento do Caixa Baixa]


Option Explicit
Dim lFlagBaixa As Integer
Dim lOpcao As String

Dim lDataEmissao As Date
Dim lNumeroConta As String
Dim lNumeroCheque As String
Dim lOrdemDigitacao As Integer
Dim lNumeroMovimentoCaixa As Long
Dim lCodigoBarra1 As String
Dim lCodigoBarra2 As String
Dim lCodigoBarra3 As String
Dim lSQL As String

'Dim rst_cheque As New adodb.Recordset
'Dim rst_baixa_cheque As New adodb.Recordset
'Dim rstTotal As New adodb.Recordset

Private BaixaCheque As New cMovimentoBaixaCheque
Private Funcionario As New cFuncionario
Private IntegracaoCaixa As New cIntegracaoCaixa
Private MovCaixa As New cMovimentoCaixa
Private MovimentoCheque As New cMovimentoCheque


Private Sub AtivaBotoes()
    cmd_alterar.Enabled = True
    cmd_extornar.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
End Sub
Private Sub Inclui()
    DesativaBotoes
    frmDados.Enabled = True
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Function IncluiMovimentoCaixa() As Boolean
    Dim xComplemento As String
    MsgBox "Recurso ainda não disponível.", vbInformation, "Módulo em Desenvolvimento"
    IncluiMovimentoCaixa = False
    lNumeroMovimentoCaixa = 0
    
    If IntegracaoCaixa.LocalizarNome(g_empresa, "DUPLICATAS A RECEBER") Then
        xComplemento = "TM:" & MovimentoCheque.TipoMovimento & " P:" & MovimentoCheque.Periodo & " " & MovimentoCheque.Emitente
        MovCaixa.Empresa = g_empresa
''''        MovCaixa.Data = Format(msk_data_baixa.Text, "dd/mm/yyyy")
        MovCaixa.NumeroMovimento = 1
        MovCaixa.Valor = MovimentoCheque.Valor
        MovCaixa.NumeroDocumento = MovimentoCheque.NumeroCheque
        MovCaixa.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
        MovCaixa.Complemento = xComplemento
        MovCaixa.NumeroContaDebito = IntegracaoCaixa.ContaDebito
        MovCaixa.NumeroContaCredito = IntegracaoCaixa.ContaCredito
        MovCaixa.TipoMovimento = 2
        MovCaixa.FluxoCaixa = True
        MovCaixa.CodigoUsuario = g_usuario
        If MovCaixa.Incluir > 0 Then
            IncluiMovimentoCaixa = True
            lNumeroMovimentoCaixa = MovCaixa.NumeroMovimento
        Else
            MsgBox "Não foi integrado no caixa o valor=" & MovimentoCheque.Valor, vbInformation, "Erro de Integridade"
        End If
    Else
        MsgBox "Não existe a integração=" & "DUPLICATAS A RECEBER" & ".", vbInformation, "Registro Inexistente"
    End If
End Function
Private Sub AtualizaMSFlexGrid()
    Dim i As Integer
    Dim rsTabela As adodb.Recordset
    
    LimpaMSFlexGrid
    lSQL = "SELECT [Data de Emissao], [Data do Vencimento], Emitente, Valor, [Numero da Conta], "
    lSQL = lSQL & "[Numero do Cheque], Periodo, Telefone, [CPF CNPJ], [Codigo do Vendedor], "
    lSQL = lSQL & "[Numero da Ilha], [Data da Custodia]"
    lSQL = lSQL & "  FROM Movimento_Cheque"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    If optEmissao.Value = True Then
        lSQL = lSQL & "   AND [Data de Emissao] >= " & preparaData(CDate(mskDataInicial.Text))
        lSQL = lSQL & "   AND [Data de Emissao] <= " & preparaData(CDate(mskDataFinal.Text))
    Else
        lSQL = lSQL & "   AND [Data do Vencimento] >= " & preparaData(CDate(mskDataInicial.Text))
        lSQL = lSQL & "   AND [Data do Vencimento] <= " & preparaData(CDate(mskDataFinal.Text))
    End If
    lSQL = lSQL & " ORDER BY Valor, [Data de Emissao], [Data do Vencimento]"
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
            MSFlexGrid.Text = rsTabela![Data de Emissao]
            MSFlexGrid.Col = 1
            MSFlexGrid.Text = rsTabela![Data do Vencimento]
            MSFlexGrid.Col = 2
            MSFlexGrid.Text = rsTabela!Emitente
            MSFlexGrid.Col = 3
            MSFlexGrid.Text = Format(rsTabela!Valor, "##,###,##0.00")
            MSFlexGrid.Col = 4
            MSFlexGrid.Text = rsTabela![Numero da Conta]
            MSFlexGrid.Col = 5
            MSFlexGrid.Text = rsTabela![Numero do Cheque]
            MSFlexGrid.Col = 6
            MSFlexGrid.Text = rsTabela![Periodo]
            MSFlexGrid.Col = 7
            MSFlexGrid.Text = fMascaraTelefone(rsTabela!Telefone)
            MSFlexGrid.Col = 8
            MSFlexGrid.Text = rsTabela![CPF CNPJ]
            MSFlexGrid.Col = 9
            MSFlexGrid.Text = rsTabela![Codigo do Vendedor]
            rsTabela.MoveNext
        Loop
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
End Sub
Private Sub AtualTabe()
    On Error GoTo FileError
    
    BaixaCheque.Empresa = g_empresa
    BaixaCheque.DataEmissao = CDate(lblDataEmissao.Caption)
    BaixaCheque.Periodo = lblPeriodo.Caption
    BaixaCheque.TipoMovimento = Val(lblTipoMovimento.Caption)
    If Len(lblCpfCnpj.Caption) = 18 Then
        BaixaCheque.CPFCNPJ = fDesmascaraCNPJ(lblCpfCnpj.Caption)
    ElseIf Len(lblCpfCnpj.Caption) = 14 Then
        BaixaCheque.CPFCNPJ = fDesmascaraCPF(lblCpfCnpj.Caption)
    Else
        BaixaCheque.CPFCNPJ = lblCpfCnpj.Caption
    End If
    BaixaCheque.Valor = fValidaValor(lblValor.Caption)
    BaixaCheque.DataVencimento = lblDataVencimento.Caption
    BaixaCheque.Agencia = lblAgencia.Caption
    BaixaCheque.BancoAgencia = lblBanco.Caption & lblAgencia.Caption
    BaixaCheque.NumeroConta = lblConta.Caption
    BaixaCheque.NumeroCheque = lblNumeroCheque.Caption
    BaixaCheque.Emitente = lblEmitente.Caption
    BaixaCheque.Telefone = fDesmascaraTelefone(lblTelefone.Caption)
    BaixaCheque.CodigoVendedor = Val(lblCodigoFuncionario.Caption)
    
    BaixaCheque.OrdemDigitacao = lOrdemDigitacao
    BaixaCheque.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
    BaixaCheque.CodigoBarra1 = lCodigoBarra1
    BaixaCheque.CodigoBarra2 = lCodigoBarra2
    BaixaCheque.CodigoBarra3 = lCodigoBarra3
    BaixaCheque.NumeroIlha = Val(lblIlha.Caption)
    If IsDate(lblDataCustodia.Caption) Then
        BaixaCheque.DataCustodia = lblDataCustodia.Caption
    Else
        BaixaCheque.DataCustodia = "00:00:00"
    End If
    BaixaCheque.DataPagamento = CDate(mskDataBaixa.Text)
    BaixaCheque.PeriodoPagamento = Val(txtPeriodoBaixa.Text)
    Exit Sub
FileError:
    MsgBox "Erro ao atualizar tabela Baixa_Cheque!", vbCritical, "Erro Interno"
    Exit Sub
End Sub
Private Sub AtualizaTabelaCheque()
    MovimentoCheque.Empresa = g_empresa
    MovimentoCheque.DataEmissao = CDate(lblDataEmissao.Caption)
    MovimentoCheque.Periodo = lblPeriodo.Caption
    MovimentoCheque.TipoMovimento = Val(lblTipoMovimento.Caption)
    If Len(lblCpfCnpj.Caption) = 18 Then
        MovimentoCheque.CPFCNPJ = fDesmascaraCNPJ(lblCpfCnpj.Caption)
    ElseIf Len(lblCpfCnpj.Caption) = 14 Then
        MovimentoCheque.CPFCNPJ = fDesmascaraCPF(lblCpfCnpj.Caption)
    Else
        MovimentoCheque.CPFCNPJ = lblCpfCnpj.Caption
    End If
    MovimentoCheque.Valor = fValidaValor(lblValor.Caption)
    MovimentoCheque.DataVencimento = lblDataVencimento.Caption
    MovimentoCheque.BancoAgencia = lblBanco.Caption & lblAgencia.Caption
    MovimentoCheque.NumeroConta = lblConta.Caption
    MovimentoCheque.NumeroCheque = lblNumeroCheque.Caption
    MovimentoCheque.Emitente = lblEmitente.Caption
    MovimentoCheque.Telefone = fDesmascaraTelefone(lblTelefone.Caption)
    MovimentoCheque.CodigoVendedor = Val(lblCodigoFuncionario.Caption)
    
    MovimentoCheque.OrdemDigitacao = lOrdemDigitacao
    MovimentoCheque.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
    MovimentoCheque.CodigoBarra1 = lCodigoBarra1
    MovimentoCheque.CodigoBarra2 = lCodigoBarra2
    MovimentoCheque.CodigoBarra3 = lCodigoBarra3
    MovimentoCheque.NumeroIlha = Val(lblIlha.Caption)
    If IsDate(lblDataCustodia.Caption) Then
        MovimentoCheque.DataCustodia = lblDataCustodia.Caption
    Else
        MovimentoCheque.DataCustodia = "00:00:00"
    End If
End Sub
Private Sub Atualtabe2()
    AtualTabe
End Sub
Private Sub AtualizaTelaCheque()
    lDataEmissao = MovimentoCheque.DataEmissao
    lNumeroConta = MovimentoCheque.NumeroConta
    lNumeroCheque = MovimentoCheque.NumeroCheque
    lOrdemDigitacao = MovimentoCheque.OrdemDigitacao
    lNumeroMovimentoCaixa = MovimentoCheque.NumeroMovimentoCaixa
    lCodigoBarra1 = MovimentoCheque.CodigoBarra1
    lCodigoBarra2 = MovimentoCheque.CodigoBarra2
    lCodigoBarra3 = MovimentoCheque.CodigoBarra3
    
    lblDataEmissao.Caption = Format(MovimentoCheque.DataEmissao, "dd/mm/yyyy")
    lblPeriodo.Caption = MovimentoCheque.Periodo
    lblIlha.Caption = MovimentoCheque.NumeroIlha
    If MovimentoCheque.TipoMovimento = 1 Then
        lblTipoMovimento.Caption = "1 - Caixa de Combustíveis"
    ElseIf MovimentoCheque.TipoMovimento = 2 Then
        lblTipoMovimento.Caption = "2 - Caixa de Óleo/Diversos"
    Else
        lblTipoMovimento.Caption = "3 - Cheque Inclusão"
    End If
    lblCpfCnpj.Caption = MovimentoCheque.CPFCNPJ
    If Len(MovimentoCheque.CPFCNPJ) = 11 Then
        lblCpfCnpj.Caption = fMascaraCPF(MovimentoCheque.CPFCNPJ)
    ElseIf Len(MovimentoCheque.CPFCNPJ) = 14 Then
        lblCpfCnpj.Caption = fMascaraCNPJ(MovimentoCheque.CPFCNPJ)
    End If
    lblValor.Caption = Format(MovimentoCheque.Valor, "###,##0.00")
    lblDataVencimento.Caption = Format(MovimentoCheque.DataVencimento, "dd/mm/yyyy")
    lblBanco.Caption = Mid(MovimentoCheque.BancoAgencia, 1, 3)
    lblAgencia.Caption = Mid(MovimentoCheque.BancoAgencia, 4, 4)
    lblConta.Caption = MovimentoCheque.NumeroConta
    lblNumeroCheque.Caption = MovimentoCheque.NumeroCheque
    lblEmitente.Caption = MovimentoCheque.Emitente
    lblTelefone.Caption = fMascaraTelefone(MovimentoCheque.Telefone)
    If MovimentoCheque.DataCustodia <> "00:00:00" Then
       lblDataCustodia.Caption = Format(MovimentoCheque.DataCustodia, "dd/mm/yyyy")
    Else
        lblDataCustodia.Caption = ""
    End If
    lblCodigoFuncionario.Caption = MovimentoCheque.CodigoVendedor
    If Funcionario.LocalizarCodigo(g_empresa, MovimentoCheque.CodigoVendedor) Then
        lblNomeFuncionario.Caption = Funcionario.Nome
    Else
        lblNomeFuncionario.Caption = "** Funcionário Não Cadastrado **"
    End If
    mskDataBaixa.Text = Format(Date, "dd/MM/yyyy")
    txtPeriodoBaixa.Text = "1"
    lOpcao = 1
End Sub
Private Sub AtualTela()
    lDataEmissao = BaixaCheque.DataEmissao
    lNumeroConta = BaixaCheque.NumeroConta
    lNumeroCheque = BaixaCheque.NumeroCheque
    lOrdemDigitacao = BaixaCheque.OrdemDigitacao
    lNumeroMovimentoCaixa = BaixaCheque.NumeroMovimentoCaixa
    lCodigoBarra1 = BaixaCheque.CodigoBarra1
    lCodigoBarra2 = BaixaCheque.CodigoBarra2
    lCodigoBarra3 = BaixaCheque.CodigoBarra3
    
    lblDataEmissao.Caption = Format(BaixaCheque.DataEmissao, "dd/mm/yyyy")
    lblPeriodo.Caption = BaixaCheque.Periodo
    lblIlha.Caption = BaixaCheque.NumeroIlha
    If BaixaCheque.TipoMovimento = 1 Then
        lblTipoMovimento.Caption = "1 - Caixa de Combustíveis"
    ElseIf BaixaCheque.TipoMovimento = 2 Then
        lblTipoMovimento.Caption = "2 - Caixa de Óleo/Diversos"
    Else
        lblTipoMovimento.Caption = "3 - Cheque Inclusão"
    End If
    lblCpfCnpj.Caption = BaixaCheque.CPFCNPJ
    If Len(BaixaCheque.CPFCNPJ) = 11 Then
        lblCpfCnpj.Caption = fMascaraCPF(BaixaCheque.CPFCNPJ)
    ElseIf Len(BaixaCheque.CPFCNPJ) = 14 Then
        lblCpfCnpj.Caption = fMascaraCNPJ(BaixaCheque.CPFCNPJ)
    End If
    lblValor.Caption = Format(BaixaCheque.Valor, "###,##0.00")
    lblDataVencimento.Caption = Format(BaixaCheque.DataVencimento, "dd/mm/yyyy")
    lblBanco.Caption = Mid(BaixaCheque.BancoAgencia, 1, 3)
    lblAgencia.Caption = BaixaCheque.Agencia
    lblConta.Caption = BaixaCheque.NumeroConta
    lblNumeroCheque.Caption = BaixaCheque.NumeroCheque
    lblEmitente.Caption = BaixaCheque.Emitente
    lblTelefone.Caption = fMascaraTelefone(BaixaCheque.Telefone)
    If BaixaCheque.DataCustodia <> "00:00:00" Then
       lblDataCustodia.Caption = Format(BaixaCheque.DataCustodia, "dd/mm/yyyy")
    Else
        lblDataCustodia.Caption = ""
    End If
    lblCodigoFuncionario.Caption = BaixaCheque.CodigoVendedor
    If Funcionario.LocalizarCodigo(g_empresa, BaixaCheque.CodigoVendedor) Then
        lblNomeFuncionario.Caption = Funcionario.Nome
    Else
        lblNomeFuncionario.Caption = "** Funcionário Não Cadastrado **"
    End If
    If BaixaCheque.DataPagamento <> "00:00:00" Then
        mskDataBaixa.Text = Format(BaixaCheque.DataPagamento, "dd/MM/yyyy")
    Else
        mskDataBaixa.Text = "__/__/____"
    End If
    txtPeriodoBaixa.Text = BaixaCheque.PeriodoPagamento
    frmDados.Enabled = False
End Sub
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
    'If Not MovCaixa.Excluir(g_empresa, lDataPagamento, lNumeroMovimentoCaixa) Then
    '    MsgBox "Não foi excluído o movimento do caixa!", vbInformation, "Erro de Integridade."
    'End If
End Sub
Private Sub Finaliza()
    Set BaixaCheque = Nothing
    Set Funcionario = Nothing
    Set IntegracaoCaixa = Nothing
    Set MovCaixa = Nothing
    Set MovimentoCheque = Nothing
End Sub
Private Sub cmd_alterar_Click()
    lOpcao = 2
    If BaixaCheque.LocalizarRegistro(g_empresa, lDataEmissao, lNumeroConta, lNumeroCheque) Then
        DesativaBotoes
        cmd_ok.Visible = True
        cmd_cancelar.Visible = True
        frmDados.Enabled = True
        'msk_data_baixa.SetFocus
        cmd_ok.SetFocus
    Else
        MsgBox "Erro de verificação."
    End If
End Sub
Private Sub cmd_anterior_Click()
    If BaixaCheque.LocalizarAnterior() Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbExclamation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    If BaixaCheque.LocalizarUltimo(g_empresa) Then
        AtivaBotoes
        AtualTela
        cmd_alterar.SetFocus
        MSFlexGrid.SetFocus
    Else
        DesativaBotoes
        cmd_sair.Enabled = True
        LimpaTela
        MSFlexGrid.SetFocus
    End If
End Sub
Private Sub LimpaTela()
    lblDataEmissao.Caption = ""
    lblPeriodo.Caption = ""
    lblIlha.Caption = ""
    lblTipoMovimento.Caption = ""
    lblCpfCnpj.Caption = ""
    lblValor.Caption = ""
    lblDataVencimento.Caption = ""
    lblBanco.Caption = ""
    lblAgencia.Caption = ""
    lblConta.Caption = ""
    lblNumeroCheque.Caption = ""
    lblEmitente.Caption = ""
    lblTelefone.Caption = ""
    lblDataCustodia.Caption = ""
    lblCodigoFuncionario.Caption = ""
    lblNomeFuncionario.Caption = ""
    mskDataBaixa.Text = "__/__/____"
    txtPeriodoBaixa.Text = ""
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
    MSFlexGrid.Text = "Data de Emissão"
    MSFlexGrid.ColWidth(i) = 1200
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Data do Vencimento"
    MSFlexGrid.ColWidth(i) = 1200
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Nome do Emitente"
    MSFlexGrid.ColWidth(i) = 2500
    MSFlexGrid.ColAlignment(i) = 1
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Valor"
    MSFlexGrid.ColWidth(i) = 1100
    MSFlexGrid.ColAlignment(i) = 7
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Número da Conta"
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = 9
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Número do Cheque"
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = 9
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Per."
    MSFlexGrid.ColWidth(i) = 400
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Telefone"
    MSFlexGrid.ColWidth(i) = 1200
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "CPF CNPJ"
    MSFlexGrid.ColWidth(i) = 1300
    MSFlexGrid.ColAlignment(i) = 9
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Codigo do Vendedor"
    MSFlexGrid.ColWidth(i) = 850
    MSFlexGrid.ColAlignment(i) = 4
    MSFlexGrid.Row = 1
    MSFlexGrid.Col = 0
End Sub
Private Sub cmd_extornar_Click()
    If BaixaCheque.LocalizarRegistro(g_empresa, lDataEmissao, lNumeroConta, lNumeroCheque) Then
        If (MsgBox("Deseja realmente extornar esta baixa?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            'If BaixaNotaAbastecimento.NumeroMovimentoCaixaBaixa > 0 Then
            '    Call ExcluiMovimentoCaixa
            'End If
            AtualizaTabelaCheque
            If MovimentoCheque.Incluir Then
                If Not BaixaCheque.Excluir(g_empresa, lDataEmissao, lNumeroConta, lNumeroCheque) Then
                    MsgBox "Não foi possível excluir este registro!", vbInformation, "Erro de Integridade"
                End If
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Integridade"
            End If
            cmd_ultimo_Click
            mskDataFinal_LostFocus
        End If
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            'If Not IncluiMovimentoCaixa Then
            '    MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
            'End If
            AtualTabe
            If BaixaCheque.Incluir Then
                If Not MovimentoCheque.Excluir(g_empresa, lDataEmissao, lNumeroConta, lNumeroCheque) Then
                    MsgBox "Não foi possível excluir este registro!", vbInformation, "Erro de Integridade"
                End If
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Integridade"
            End If
        ElseIf lOpcao = 2 Then
            Atualtabe2
            If Not BaixaCheque.Alterar(g_empresa, lDataEmissao, lNumeroConta, lNumeroCheque) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Integridade"
            End If
        End If
        AtualizaMSFlexGrid
        If BaixaCheque.LocalizarRegistro(g_empresa, lDataEmissao, lNumeroConta, lNumeroCheque) Then
            AtualTela
        Else
            MsgBox "Não foi possível localizar o registro!", vbInformation, "Erro de Integridade"
        End If
        MSFlexGrid.SetFocus
    End If
    Exit Sub
FileError:
    'ErroArquivo tbl_movimento_nota.Name, "Notaa"
    Exit Sub
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(mskDataBaixa.Text) Then
        MsgBox "Informe a data da baixa.", vbInformation, "Atenção!"
        mskDataBaixa.SetFocus
    ElseIf CDate(mskDataBaixa.Text) < CDate(lblDataEmissao.Caption) Then
        MsgBox "A data da baixa deve ser maior ou igual a " & lblDataEmissao.Caption & ".", vbInformation, "Atenção!"
        mskDataBaixa.SetFocus
    ElseIf Val(txtPeriodoBaixa.Text) = 0 Then
        MsgBox "Informe o período da baixado.", vbInformation, "Atenção!"
        txtPeriodoBaixa.SetFocus
'    ElseIf fValidaValor2(txt_valor_baixado.Text) < fValidaValor2(lbl_total.Caption) Then
'        MsgBox "O valor baixado não pode ser menor que " & lbl_total.Caption & ".", vbInformation, "Atenção!"
'        txt_valor_baixado.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_cheque_baixado.Show 1
    If Len(g_string) > 0 Then
        lDataEmissao = RetiraGString(1)
        lNumeroConta = RetiraGString(2)
        lNumeroCheque = RetiraGString(3)
        If BaixaCheque.LocalizarRegistro(g_empresa, lDataEmissao, lNumeroConta, lNumeroCheque) Then
            AtualTela
            cmd_alterar.SetFocus
        Else
            LimpaTela
            MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If BaixaCheque.LocalizarPrimeiro(g_empresa) Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Cliente não tem baixa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If BaixaCheque.LocalizarProximo() Then
        AtualTela
    Else
        MsgBox "Fim de Arquivo.", vbExclamation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Finaliza
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If BaixaCheque.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não existe cheque baixado a ser mostrado!.", vbInformation, "Erro de Verificação!"
        LimpaTela
    End If
End Sub
Private Sub cmdDataFinal_Click()
    g_string = mskDataFinal.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        mskDataInicial.Text = RetiraGString(1)
        mskDataFinal.Text = RetiraGString(2)
    Else
        mskDataFinal.Text = RetiraGString(1)
    End If
    g_string = ""
    MSFlexGrid.SetFocus
End Sub
Private Sub cmdDataInicial_Click()
    g_string = mskDataInicial.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        mskDataInicial.Text = RetiraGString(1)
        mskDataFinal.Text = RetiraGString(2)
        MSFlexGrid.SetFocus
    Else
        mskDataInicial.Text = RetiraGString(1)
        mskDataFinal.SetFocus
    End If
    g_string = ""
End Sub
Private Sub Form_Activate()
    If lFlagBaixa = 0 Then
        lFlagBaixa = 1
        lOpcao = 0
        DesativaBotoes
        mskDataInicial.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
        mskDataFinal.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
        AtualizaMSFlexGrid
        If BaixaCheque.LocalizarUltimo(g_empresa) Then
            AtualTela
            AtivaBotoes
        End If
        MSFlexGrid.SetFocus
    End If
    Screen.MousePointer = 1
End Sub
Private Sub MarcaCelulas()
    If MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0) <> "" Then
        lDataEmissao = MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0)
        lNumeroConta = MSFlexGrid.TextMatrix(MSFlexGrid.Row, 4)
        lNumeroCheque = MSFlexGrid.TextMatrix(MSFlexGrid.Row, 5)
        If MovimentoCheque.LocalizarRegistro(g_empresa, lDataEmissao, lNumeroConta, lNumeroCheque) Then
            Inclui
            AtualizaTelaCheque
            mskDataBaixa.SetFocus
        Else
        End If
    End If
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
Private Sub Form_Load()
    CentraForm Me
End Sub
Private Sub MSFlexGrid_DblClick()
    MarcaCelulas
End Sub
Private Sub MSFlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        MarcaCelulas
    ElseIf KeyCode = 98 Then
        KeyCode = 40
    ElseIf KeyCode = 99 Then
        KeyCode = 34
    ElseIf KeyCode = 104 Then
        KeyCode = 38
    ElseIf KeyCode = 105 Then
        KeyCode = 33
    End If
End Sub
Private Sub MSFlexGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0
        MarcaCelulas
    End If
End Sub
Private Sub mskDataBaixa_GotFocus()
    mskDataBaixa.SelStart = 0
    mskDataBaixa.SelLength = 2
End Sub
Private Sub mskDataBaixa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtPeriodoBaixa.SetFocus
    End If
End Sub
Private Sub mskDataFinal_GotFocus()
    mskDataFinal.SelStart = 0
    mskDataFinal.SelLength = 2
End Sub
Private Sub mskDataFinal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        MSFlexGrid.SetFocus
    End If
End Sub
Private Sub mskDataFinal_LostFocus()
    AtualizaMSFlexGrid
End Sub
Private Sub mskDataInicial_GotFocus()
    mskDataInicial.SelStart = 0
    mskDataInicial.SelLength = 2
End Sub
Private Sub mskDataInicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        mskDataFinal.SetFocus
    End If
End Sub
Private Sub txtPeriodoBaixa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
