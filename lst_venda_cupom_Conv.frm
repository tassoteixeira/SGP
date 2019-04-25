VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_venda_cupom_Conv 
   Caption         =   "Emissão das Vendas de Cupom Fiscal - Conveniência"
   ClientHeight    =   3810
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6900
   Icon            =   "lst_venda_cupom_Conv.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_venda_cupom_Conv.frx":030A
   ScaleHeight     =   3810
   ScaleWidth      =   6900
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   975
      Left            =   1140
      Picture         =   "lst_venda_cupom_Conv.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Visualiza inventário de produtos."
      Top             =   2700
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   975
      Left            =   3000
      Picture         =   "lst_venda_cupom_Conv.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprime inventário de produtos."
      Top             =   2700
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   975
      Left            =   4860
      Picture         =   "lst_venda_cupom_Conv.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2700
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2475
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.ComboBox cboTipoVenda 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1500
         Width           =   4755
      End
      Begin VB.ComboBox cbo_periodo_i 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
         Width           =   495
      End
      Begin VB.ComboBox cbo_periodo_f 
         Height          =   315
         Left            =   4860
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_venda_cupom_Conv.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "lst_venda_cupom_Conv.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_venda_cupom_Conv.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
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
      Begin VB.Label Label6 
         Caption         =   "&Tipo de Venda"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Período inicial"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "Período &final"
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "&Data de Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_venda_cupom_Conv"
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

Dim lSubQuantidade As Currency
Dim lSubValor As Currency
Dim lSubTotalQuantidade As Currency
Dim lSubTotalValor As Currency
Dim lTotalQuantidade As Currency
Dim lTotalValor As Currency
Dim lImprimirECF As Boolean
Dim BemaRetorno As Integer
Dim lNomeGrupo As String
Dim lSQL As String
Dim lSerieECF As String
Dim lEcfInstalada As Boolean
Dim lUtilizaNFCe As Boolean 'Define se irei buscar os dados da tabela cupom ou documento eletronico

Dim lDataDetalhada As Date

Dim lImpBematech As Boolean
Dim lImpSchalter As Boolean
Dim lImpMecaf As Boolean
Dim lImpQuick As Boolean
Dim lImpDaruma As Boolean

Dim lLocalImpressaoTermica As Integer

Dim rs As New adodb.Recordset
Dim rstVendaCupomFiscal As New adodb.Recordset

Private Aliquota As New CadastroDLL.cAliquota
Private ECF As New CadastroDLL.cEcf
Private MovimentoCupomFiscal As New CadastroDLL.cMovimentoCupomFiscal
Private MovimentoCupomFiscalItem As New CadastroDLL.cMovimentoCupomFiscalItem

Private ConfiguracaoDiversa As New CadastroDLL.cConfiguracaoDiversa

Private MovDocumentoEletronicoCabecalho As New CadastroDLL.cMovDocEletronicoCabecalho
Private MovDocumentoEletronicoItem As New CadastroDLL.cMovDocEletronicoItem

Private Produto As New CadastroDLL.cProduto

Private Sub AtualizaConstantes()
    Dim dados As String
    
    lEcfInstalada = False
    dados = ReadINI("CUPOM FISCAL", "ECF Instalada", gArquivoIni)
    If dados = "SIM" Then
        lEcfInstalada = True
    End If
    
    lUtilizaNFCe = False
    If ConfiguracaoDiversa.LocalizarCodigo(g_empresa, "NFCe: Numero") = True Then
       If ConfiguracaoDiversa.Codigo > 0 Then
          lUtilizaNFCe = True
          Me.Caption = Me.Caption & " - NFCe "
          cbo_periodo_i.ListIndex = 0
          cbo_periodo_f.ListIndex = cbo_periodo_f.ListCount - 1
          cbo_periodo_i.Enabled = False
          cbo_periodo_f.Enabled = False
       End If
    End If

    
    lImpBematech = False
    lImpSchalter = False
    lImpMecaf = False
    lImpQuick = False
    lImpDaruma = False
    dados = ReadINI("CUPOM FISCAL", "Impressora Fiscal", gArquivoIni)
    Call CriaLogCupom("AtualizaConstantes - Marca da Impressora Fiscal=" & dados)
    
    If Not lUtilizaNFCe Then
    
        Me.Caption = Me.Caption & " - ECF: " & dados
        If dados = "BEMATECH" Then
            lImpBematech = True
        ElseIf dados = "SCHALTER" Then
            lImpSchalter = True
        ElseIf dados = "MECAF" Then
            lImpMecaf = True
        ElseIf dados = "QUICK" Then
            lImpQuick = True
        ElseIf dados = "DARUMA" Then
            lImpDaruma = True
        End If
        
    End If
    
    
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Aliquota = Nothing
    Set ECF = Nothing
    Set MovimentoCupomFiscal = Nothing
    Set MovimentoCupomFiscalItem = Nothing
    Set Produto = Nothing
    Set rstVendaCupomFiscal = Nothing
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lSubQuantidade = 0
    lSubValor = 0
    lSubTotalQuantidade = 0
    lSubTotalValor = 0
    lTotalQuantidade = 0
    lTotalValor = 0
    lSerieECF = "12345678901234"
    
    If ECF.LocalizarUltimo(g_empresa) Then
        lSerieECF = ECF.NumeroSerie
    Else
    End If
End Sub
Private Sub ZzLoopRecalculaDescontoCupom()
    Dim xTotalCupom As Currency
    Dim xTotalRecebido As Currency
    Dim xTotalDesconto As Currency
    Dim xCodigoEcf As Integer
    Dim xUltimaOrdem As Integer
    Dim rstNumCupom As New adodb.Recordset
    Dim rstMovCupom As New adodb.Recordset
        
        
    On Error GoTo FileError
    
    
    lSQL = ""
    lSQL = lSQL & " SELECT Data, [Numero do Cupom]"
    lSQL = lSQL & "   FROM Movimento_Cupom_Fiscal"
    lSQL = lSQL & "  WHERE Empresa = " & g_empresa
    lSQL = lSQL & "    AND Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "    AND Data <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & "    AND Periodo >= " & preparaTexto(cbo_periodo_i.Text)
    lSQL = lSQL & "    AND Periodo <= " & preparaTexto(cbo_periodo_f.Text)
    lSQL = lSQL & "    AND [Cupom Cancelado] = " & preparaBooleano(False)
    lSQL = lSQL & "    AND [Item Cancelado] = " & preparaBooleano(False)
    lSQL = lSQL & "  GROUP BY Data, [Numero do Cupom]"
    lSQL = lSQL & "  ORDER BY Data, [Numero do Cupom]"
    Set rstNumCupom = Conectar.RsConexao(lSQL)
    If rstNumCupom.RecordCount > 0 Then
        Do Until rstNumCupom.EOF
            lSQL = ""
            lSQL = lSQL & " SELECT [Valor Total], [Valor Recebido], [Valor do Desconto], Ordem, [Codigo da Ecf]"
            lSQL = lSQL & "   FROM Movimento_Cupom_Fiscal"
            lSQL = lSQL & "  WHERE Empresa = " & g_empresa
            lSQL = lSQL & "    AND [Numero do Cupom] = " & rstNumCupom("Numero do Cupom").Value
            lSQL = lSQL & "    AND Data = " & preparaData(rstNumCupom("Data").Value)
            lSQL = lSQL & "  ORDER BY Ordem"
            xTotalCupom = 0
            xTotalRecebido = 0
            xTotalDesconto = 0
            Set rstMovCupom = Conectar.RsConexao(lSQL)
            If rstMovCupom.RecordCount > 0 Then
                Do Until rstMovCupom.EOF
                    xTotalCupom = xTotalCupom + rstMovCupom("Valor Total").Value
                    xTotalRecebido = rstMovCupom("Valor Recebido").Value
                    xTotalDesconto = xTotalDesconto + rstMovCupom("Valor do Desconto").Value
    
                    xCodigoEcf = rstMovCupom("Codigo da Ecf").Value
                    xUltimaOrdem = rstMovCupom("Ordem").Value
                    
                    rstMovCupom.MoveNext
                Loop
                If xTotalCupom <> (xTotalRecebido + xTotalDesconto) Then
                    'MsgBox "Cupom = " & rstNumCupom("Numero do Cupom").Value & " - Data = " & rstNumCupom("Data").Value
                    If xTotalCupom = xTotalRecebido Then
                        
                        'MsgBox "Desconsiderar o desconto = " & xTotalDesconto & vbCrLf & "Cupom = " & rstNumCupom("Numero do Cupom").Value & " - Data = " & rstNumCupom("Data").Value
                        'If xUltimaOrdem = 1 Then
                            If MovimentoCupomFiscal.LocalizarCodigo(g_empresa, xCodigoEcf, rstNumCupom("Numero do Cupom").Value, rstNumCupom("Data").Value, 1) Then
                                MovimentoCupomFiscal.ValorDesconto = 0
                                If Not MovimentoCupomFiscal.Alterar(g_empresa, xCodigoEcf, rstNumCupom("Numero do Cupom").Value, rstNumCupom("Data").Value, 1) Then
                                    MsgBox "erro ao alterar o cupom=" & rstNumCupom("Numero do Cupom").Value & " - Data = " & rstNumCupom("Data").Value
                                End If
                            End If
                        'End If
                    End If
                End If
            End If
            rstMovCupom.Close
            rstNumCupom.MoveNext
        Loop
        MsgBox "Processamento concluído"
    End If
    rstNumCupom.Close
    Set rstMovCupom = Nothing
    Set rstNumCupom = Nothing
    Exit Sub

FileError:
    MsgBox " - ZzLoopRecalculaDescontoCupom: Erro ao processar recalculo de desconto de cupom fiscal"
    Exit Sub
End Sub
'Private Sub PreencheCboFormaPagamento()
'    cboFormaPagamento.Clear
'    cboFormaPagamento.AddItem "0 - Todas"
'    cboFormaPagamento.ItemData(cboFormaPagamento.NewIndex) = 0
'    cboFormaPagamento.AddItem "1 - Dinheiro"
'    cboFormaPagamento.ItemData(cboFormaPagamento.NewIndex) = 1
'    cboFormaPagamento.AddItem "2 - Cheque à Vista"
'    cboFormaPagamento.ItemData(cboFormaPagamento.NewIndex) = 2
'    cboFormaPagamento.AddItem "3 - Cheque Pré-Datado"
'    cboFormaPagamento.ItemData(cboFormaPagamento.NewIndex) = 3
'    cboFormaPagamento.AddItem "4 - Cartão de Crédito"
'    cboFormaPagamento.ItemData(cboFormaPagamento.NewIndex) = 4
'    cboFormaPagamento.AddItem "5 - Nota Vinculada"
'    cboFormaPagamento.ItemData(cboFormaPagamento.NewIndex) = 5
'    cboFormaPagamento.AddItem "6 - Cartão TecBan    "
'    cboFormaPagamento.ItemData(cboFormaPagamento.NewIndex) = 6
'    cboFormaPagamento.AddItem "7 - Cheque TecBan    "
'    cboFormaPagamento.ItemData(cboFormaPagamento.NewIndex) = 7
'    cboFormaPagamento.AddItem "8 - Ticket Car Smart "
'    cboFormaPagamento.ItemData(cboFormaPagamento.NewIndex) = 8
'    cboFormaPagamento.AddItem "9 - Smart Shop/Check Check"
'    cboFormaPagamento.ItemData(cboFormaPagamento.NewIndex) = 9
'    cboFormaPagamento.AddItem "10 - SuperCard"
'    cboFormaPagamento.ItemData(cboFormaPagamento.NewIndex) = 10
'End Sub
'Private Sub PreencheCboFuncionario()
'    'Prepara SQL
'    lSQL = ""
'    lSQL = lSQL & "   SELECT Codigo, Nome"
'    lSQL = lSQL & "     FROM Funcionario"
'    lSQL = lSQL & "    WHERE Empresa = " & g_empresa
'    lSQL = lSQL & "      AND Situacao = " & preparaTexto("A")
'    lSQL = lSQL & "      AND Periodo < " & 5
'    lSQL = lSQL & " ORDER BY Nome, Codigo"
'    'Abre RecordSet
'    Set rs = Conectar.RsConexao(lSQL)
'
'    cboFuncionario.Clear
'    cboFuncionario.AddItem "Todos os Funcionários"
'    cboFuncionario.ItemData(cboFuncionario.NewIndex) = 0
'    If rs.RecordCount > 0 Then
'        rs.MoveFirst
'        Do Until rs.EOF
'            cboFuncionario.AddItem rs("Nome").Value
'            cboFuncionario.ItemData(cboFuncionario.NewIndex) = rs("Codigo").Value
'            rs.MoveNext
'        Loop
'    End If
'    Set rs = Nothing
'End Sub

'Private Sub PreencheCboProduto()
'    'Prepara SQL
'    lSQL = ""
'    lSQL = lSQL & "   SELECT Codigo, Nome"
'    lSQL = lSQL & "     FROM Produto"
'    lSQL = lSQL & "    WHERE Inativo = " & 0
'    'lSQL = lSQL & "      AND Inativo = " & 0
'
'    lSQL = lSQL & " ORDER BY Nome, Codigo"
'    'Abre RecordSet
'    Set rs = Conectar.RsConexao(lSQL)
'
'    cboProduto.Clear
'    cboProduto.AddItem "Todos os Produtos"
'    cboProduto.ItemData(cboProduto.NewIndex) = 0
'    If rs.RecordCount > 0 Then
'        rs.MoveFirst
'        Do Until rs.EOF
'            cboProduto.AddItem rs("Nome").Value
'            cboProduto.ItemData(cboProduto.NewIndex) = rs("Codigo").Value
'            rs.MoveNext
'        Loop
'    End If
'    Set rs = Nothing
'End Sub

'Private Sub PreencheCboCliente()
'    'Prepara SQL
'    lSQL = ""
'    lSQL = lSQL & "   SELECT Codigo, [Razao Social]"
'    lSQL = lSQL & "     FROM Cliente"
'    lSQL = lSQL & "    WHERE Empresa = " & g_empresa
'    lSQL = lSQL & "      AND Inativo = " & 0
'
'    'lSQL = lSQL & "      AND Inativo = " & 0
'
'    lSQL = lSQL & " ORDER BY [Razao Social], Codigo"
'    'Abre RecordSet
'    Set rs = Conectar.RsConexao(lSQL)
'
'    cboCliente.Clear
'    cboCliente.AddItem "Todos os Clientes"
'    cboCliente.ItemData(cboCliente.NewIndex) = 0
'    If rs.RecordCount > 0 Then
'        rs.MoveFirst
'        Do Until rs.EOF
'            cboCliente.AddItem rs("Razao Social").Value
'            cboCliente.ItemData(cboCliente.NewIndex) = rs("Codigo").Value
'            rs.MoveNext
'        Loop
'    End If
'    Set rs = Nothing
'End Sub
'Private Sub PreencheCboGrupoProduto()
'    'Prepara SQL
'    lSQL = ""
'    lSQL = lSQL & "   SELECT Codigo, Nome"
'    lSQL = lSQL & "     FROM Grupo"
'    'lSQL = lSQL & "    WHERE Empresa = " & g_empresa
'    'lSQL = lSQL & "      AND Inativo = " & 0
'
'    'lSQL = lSQL & "      AND Inativo = " & 0
'
'    lSQL = lSQL & " ORDER BY Nome, Codigo"
'    'Abre RecordSet
'    Set rs = Conectar.RsConexao(lSQL)
'
'    cboGrupoProduto.Clear
'    cboGrupoProduto.AddItem "Todos os Grupos"
'    cboGrupoProduto.ItemData(cboGrupoProduto.NewIndex) = 0
'    If rs.RecordCount > 0 Then
'        rs.MoveFirst
'        Do Until rs.EOF
'            cboGrupoProduto.AddItem rs("Nome").Value
'            cboGrupoProduto.ItemData(cboGrupoProduto.NewIndex) = rs("Codigo").Value
'            rs.MoveNext
'        Loop
'    End If
'    Set rs = Nothing
'End Sub
'Private Sub PreencheCboTributacao()
'    'Prepara SQL
'    lSQL = ""
'    lSQL = lSQL & "   SELECT Codigo, Nome"
'    lSQL = lSQL & "     FROM Aliquota"
'    'lSQL = lSQL & "    WHERE Empresa = " & g_empresa
'    'lSQL = lSQL & "      AND Inativo = " & 0
'
'    'lSQL = lSQL & "      AND Inativo = " & 0
'    lSQL = lSQL & " GROUP BY Nome, Codigo"
'    lSQL = lSQL & " ORDER BY Nome, Codigo"
'    'Abre RecordSet
'    Set rs = Conectar.RsConexao(lSQL)
'
'    cboTributacao.Clear
'    cboTributacao.AddItem "Todas as Tributações"
'    cboTributacao.ItemData(cboTributacao.NewIndex) = 0
'    If rs.RecordCount > 0 Then
'        rs.MoveFirst
'        Do Until rs.EOF
'            cboTributacao.AddItem rs("Nome").Value
'            cboTributacao.ItemData(cboTributacao.NewIndex) = rs("Codigo").Value
'            rs.MoveNext
'        Loop
'    End If
'    Set rs = Nothing
'End Sub
Private Sub PreencheCboPeriodo()
    cbo_periodo_i.Clear
    cbo_periodo_f.Clear
    cbo_periodo_i.AddItem 1
    cbo_periodo_f.AddItem 1
    cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = 1
    cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = 1
    cbo_periodo_i.AddItem 2
    cbo_periodo_f.AddItem 2
    cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = 2
    cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = 2
    cbo_periodo_i.AddItem 3
    cbo_periodo_f.AddItem 3
    cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = 3
    cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = 3
    cbo_periodo_i.AddItem 4
    cbo_periodo_f.AddItem 4
    cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = 4
    cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = 4
    cbo_periodo_i.AddItem 5
    cbo_periodo_f.AddItem 5
    cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = 5
    cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = 5
End Sub
Private Sub PreencheCboTipoVenda()
    cboTipoVenda.Clear
    cboTipoVenda.AddItem "1 - Vendas do Posto"
    cboTipoVenda.ItemData(cboTipoVenda.NewIndex) = 1
    cboTipoVenda.AddItem "2 - Vendas da Conveniência"
    cboTipoVenda.ItemData(cboTipoVenda.NewIndex) = 2
    cboTipoVenda.AddItem "3 - Vendas Geral"
    cboTipoVenda.ItemData(cboTipoVenda.NewIndex) = 3
End Sub
Private Sub LoopRelatorio()
    ZeraVariaveis
'    If chkDetalhadaData.Value = 1 Then
'        lDataDetalhada = CDate(msk_data_i.Text)
'        Do Until lDataDetalhada > CDate(msk_data_f.Text)
'            Relatorio
'            lDataDetalhada = DateAdd("d", 1, lDataDetalhada)
'        Loop
'    Else
If Not lUtilizaNFCe Then
    Relatorio
Else
    RelatorioNFCe
End If
        
'    End If
    cmd_sair.SetFocus
End Sub
Private Sub Relatorio()
    
    'Seleciona Produtos Vendidos dentro das condições
    lSQL = ""
    lSQL = lSQL & "SELECT "
    lSQL = lSQL & "Grupo.Nome AS NomeGrupo, Produto.Nome AS NomeProduto, Produto.Unidade, "
    lSQL = lSQL & "Movimento_Cupom_Fiscal.[Codigo do Grupo], Movimento_Cupom_Fiscal.[Codigo do Produto], "
    lSQL = lSQL & "Movimento_Cupom_Fiscal.[Codigo da Aliquota], "
    lSQL = lSQL & "SUM(Quantidade) AS TotQuantidade, "
    'lSQL = lSQL & "SUM([Valor Total]) AS TotTotal" 'old 29/10/15
    lSQL = lSQL & "SUM(Movimento_Cupom_Fiscal.[Valor Total] - Movimento_Cupom_Fiscal.[Valor do Desconto]) As TotTotal" 'new 29/10/15
    lSQL = lSQL & "  FROM movimento_cupom_fiscal, Produto, Grupo"
    lSQL = lSQL & " WHERE movimento_cupom_fiscal.Empresa = " & g_empresa
    
''    If chkDetalhadaData.Value = 1 Then
''       ' lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.Data >= " & preparaData(lDataDetalhada)
''        lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.Data = " & preparaData(lDataDetalhada)
''    Else
''        lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.Data >= " & preparaData(msk_data_i.Text)
''        lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.Data <= " & preparaData(msk_data_f.Text)
''    End If

    lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.Data <= " & preparaData(msk_data_f.Text)

    
    lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.Periodo >= " & preparaTexto(cbo_periodo_i.Text)
    lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.Periodo <= " & preparaTexto(cbo_periodo_f.Text)
    lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.[Codigo do Produto] = Produto.Codigo"
    lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.[Codigo do Grupo] = Grupo.Codigo"
    

''    If cboFuncionario.ListIndex > 0 Then
''        lSQL = lSQL & " AND Operador = " & Val(cboFuncionario.ItemData(cboFuncionario.ListIndex))
''    End If
''    If cboFormaPagamento.ListIndex > 0 Then
''        lSQL = lSQL & " AND [Forma de Pagamento] = " & Val(cboFormaPagamento.ItemData(cboFormaPagamento.ListIndex))
''    End If
    If cboTipoVenda.ListIndex = 0 Then
        lSQL = lSQL & " AND Produto.[Exclusivo Posto] = " & preparaBooleano(True)
    End If
    If cboTipoVenda.ListIndex = 1 Then
        lSQL = lSQL & " AND Produto.[Exclusivo Loja] = " & preparaBooleano(True)
    End If
''    If cboProduto.ListIndex > 0 Then
''        lSQL = lSQL & " AND [Codigo do Produto] = " & Val(cboProduto.ItemData(cboProduto.ListIndex))
''    End If
''    If cboCliente.ListIndex > 0 Then
''        lSQL = lSQL & " AND [Codigo do Cliente] = " & Val(cboCliente.ItemData(cboCliente.ListIndex))
''    End If
''    If cboGrupoProduto.ListIndex > 0 Then
''        lSQL = lSQL & " AND Movimento_Cupom_Fiscal.[Codigo do Grupo] = " & Val(cboGrupoProduto.ItemData(cboGrupoProduto.ListIndex))
''    End If
''    If cboTributacao.ListIndex > 0 Then
''        lSQL = lSQL & " AND Movimento_Cupom_Fiscal.[Codigo da Aliquota] = " & Val(cboGrupoProduto.ItemData(cboTributacao.ListIndex))
''    End If
    
''    If chkSomenteCancelado.Value = 1 Then
''        lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.[Cupom Cancelado] = " & preparaBooleano(True)
''        lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.[Item Cancelado] = " & preparaBooleano(True)
''    Else
''        lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.[Cupom Cancelado] = " & preparaBooleano(False)
''        lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.[Item Cancelado] = " & preparaBooleano(False)
''    End If
    
    lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.[Cupom Cancelado] = " & preparaBooleano(False)
    lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.[Item Cancelado] = " & preparaBooleano(False)

    
    lSQL = lSQL & " GROUP BY  Produto.Nome, Produto.Unidade, Grupo.Nome, "
    lSQL = lSQL & " Movimento_Cupom_Fiscal.[Codigo do Grupo], Movimento_Cupom_Fiscal.[Codigo do Produto], movimento_cupom_fiscal.[Codigo da Aliquota]"
    lSQL = lSQL & " ORDER BY  NomeGrupo, NomeProduto, Produto.Unidade, "
    lSQL = lSQL & " Movimento_Cupom_Fiscal.[Codigo do Grupo], Movimento_Cupom_Fiscal.[Codigo do Produto], movimento_cupom_fiscal.[Codigo da Aliquota]"
    
    
    Set rstVendaCupomFiscal = Conectar.RsConexao(lSQL)
    If rstVendaCupomFiscal.RecordCount > 0 Then
        ImpDados
''    ElseIf chkDetalhadaData.Value = 1 Then 'alex
''       If lDataDetalhada = CDate(msk_data_f.Text) Then
''          FinalizaRelatorioDataDetalhada
''       End If
    End If
    rstVendaCupomFiscal.Close
End Sub
Private Sub RelatorioNFCe()
    
    'Seleciona Produtos Vendidos dentro das condições
    lSQL = ""
    lSQL = lSQL & "SELECT "
    lSQL = lSQL & "(Grupo.Nome + ' - CFOP: ' + Grupo.[CFOP de Saida]) AS NomeGrupo, Produto.Nome AS NomeProduto, Produto.Unidade, "
    lSQL = lSQL & "Produto.[Codigo do Grupo], IdProduto_MovDEItem AS [Codigo do Produto], "
    lSQL = lSQL & "Produto.[Codigo da Aliquota], "
    lSQL = lSQL & "SUM(Quantidade_MovDEItem) AS TotQuantidade, "
    lSQL = lSQL & "SUM(ValorTotalLiquido_MovDEItem) AS TotTotal"
    lSQL = lSQL & "  FROM MovimentoDocumentoEletronicoItem, Produto, Grupo"
    lSQL = lSQL & " WHERE IdEstabelecimento_MovDEItem = " & g_empresa
    lSQL = lSQL & "   AND DataEmissao_MovDEItem >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND DataEmissao_MovDEItem <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & "   AND IdProduto_MovDEItem = Produto.Codigo"
    lSQL = lSQL & "   AND Produto.[Codigo do Grupo] = Grupo.Codigo"
    If cboTipoVenda.ListIndex = 0 Then
        lSQL = lSQL & " AND Produto.[Exclusivo Posto] = " & preparaBooleano(True)
    End If
    If cboTipoVenda.ListIndex = 1 Then
        lSQL = lSQL & " AND Produto.[Exclusivo Loja] = " & preparaBooleano(True)
    End If
    lSQL = lSQL & "   AND Cancelado_MovDEItem = " & preparaBooleano(False)
    lSQL = lSQL & "   AND EtapaConcluida_MovDEItem = 9"
    
    lSQL = lSQL & " GROUP BY  Produto.Nome, Produto.Unidade, Grupo.Nome, Grupo.[CFOP de Saida],"
    lSQL = lSQL & " Produto.[Codigo do Grupo], IdProduto_MovDEItem, Produto.[Codigo da Aliquota]"
    lSQL = lSQL & " ORDER BY  NomeGrupo, NomeProduto, Produto.Unidade, "
    lSQL = lSQL & " Produto.[Codigo do Grupo], IdProduto_MovDEItem, Produto.[Codigo da Aliquota]"
    
    Set rstVendaCupomFiscal = Conectar.RsConexao(lSQL)
    
    If rstVendaCupomFiscal.RecordCount > 0 Then
        ImpDados
    End If
    rstVendaCupomFiscal.Close
End Sub
Private Sub FinalizaRelatorioDataDetalhada() 'alex
    If lPagina > 0 Then
        If lSubTotalValor > 0 Then
            ImpSubTotal
            lSubTotalQuantidade = 0
            lSubTotalValor = 0
        End If
        ImpTotal
        If lImprimirECF = False Then
            BioImprime "@@Printer.FontName = Courier New"
            BioImprime "@Printer.Print " & " "
            BioImprime "@@Printer.EndDoc"
            BioFechaImprime
            g_string = lLocal & lNomeArquivo & "|@|Relatório da Venda com Cupom Fiscal|@|"
            frm_preview.Show 1
        Else
            If fUsaNFCe Then 'alex - termica
                Call ImpTermicaRodape
                Call ImpTermicaFechaRelatorio("Relatório da Venda com Cupom Fiscal")
            ElseIf lImpBematech Then
                'Fechamento de Relatório Gerencial
                BemaRetorno = Bematech_FI_FechaRelatorioGerencial
            ElseIf lImpQuick Then
                'Fecha Relatorio Gerencial
                If EcfQuickEncerraDocumento("", "Cerrado Informatica") Then
                    BemaRetorno = 1
                Else
                    BemaRetorno = -1
                End If
            ElseIf lImpDaruma Then
                Call CriaLogCupom("ImpDados - Daruma_TEF_FechaRelatorio.")
                BemaRetorno = Daruma_TEF_FechaRelatorio()
                Call CriaLogCupom("ImpDados - Daruma_TEF_FechaRelatorio. BemaRetorno=" & BemaRetorno)
            End If
        End If
     End If
End Sub
Private Sub ImpDados()
    LoopTabelaGrupo
    If lPagina > 0 Then
        If lSubTotalValor > 0 Then
            ImpSubTotal
            lSubTotalQuantidade = 0
            lSubTotalValor = 0
        End If
        ImpTotal
        
''        If chkDetalhadaData.Value = 1 Then
''            If lDataDetalhada = CDate(msk_data_f.Text) Then
''                FinalizaRelatorioDataDetalhada 'alex
''            End If
''        Else
            If lImprimirECF = False Then
                BioImprime "@@Printer.FontName = Courier New"
                BioImprime "@Printer.Print " & " "
                BioImprime "@@Printer.EndDoc"
                BioFechaImprime
                g_string = lLocal & lNomeArquivo & "|@|Relatório da Venda com Cupom Fiscal|@|"
                frm_preview.Show 1
            Else
            
                If fUsaNFCe Then 'alex - termica
                        Call ImpTermicaRodape
                        Call ImpTermicaFechaRelatorio("Relatório da Venda com Cupom Fiscal")
                ElseIf lImpBematech Then
                    'Fechamento de Relatório Gerencial
                    BemaRetorno = Bematech_FI_FechaRelatorioGerencial
                ElseIf lImpQuick Then
                    'Fecha Relatorio Gerencial
                    If EcfQuickEncerraDocumento("", "Cerrado Informatica") Then
                        BemaRetorno = 1
                    Else
                        BemaRetorno = -1
                    End If
                ElseIf lImpDaruma Then
                    Call CriaLogCupom("ImpDados - Daruma_TEF_FechaRelatorio.")
                    BemaRetorno = Daruma_TEF_FechaRelatorio()
                    Call CriaLogCupom("ImpDados - Daruma_TEF_FechaRelatorio. BemaRetorno=" & BemaRetorno)
                End If
            End If
''        End If
    End If
End Sub
Private Sub LoopTabelaGrupo()
    'loop tabela
    Do Until rstVendaCupomFiscal.EOF
        If lNomeGrupo <> rstVendaCupomFiscal!NomeGrupo Then
            If lSubTotalValor > 0 Then
                ImpSubTotal
                lSubTotalQuantidade = 0
                lSubTotalValor = 0
            End If
            lNomeGrupo = rstVendaCupomFiscal!NomeGrupo
        End If
        If lImprimirECF = True And rstVendaCupomFiscal![Codigo do Grupo] = 4 Then
        
        Else
            LoopTabelaMovimentoCupomFiscal
        End If
        rstVendaCupomFiscal.MoveNext
    Loop
End Sub
Private Sub LoopTabelaMovimentoCupomFiscal()
''''    If chkImprimeDetalhe.Value = 1 And fUsaNFCe = False Then
''        'Seleciona Produtos Vendidos dentro das condições
''        lSQL = ""
''        lSQL = lSQL & " SELECT Produto.Nome, Produto.Unidade, Movimento_Cupom_Fiscal.Operador, Movimento_Cupom_Fiscal.Quantidade, Movimento_Cupom_Fiscal.[Valor Unitario], Movimento_Cupom_Fiscal.[Valor Total], Movimento_Cupom_Fiscal.[Valor do Desconto], Movimento_Cupom_Fiscal.Data, Movimento_Cupom_Fiscal.[Numero do Cupom], Movimento_Cupom_Fiscal.Ordem, Movimento_Cupom_Fiscal.[Tipo do SubEstoque], Funcionario.Nome AS NomeFuncionario"
''        lSQL = lSQL & "   FROM Produto, Movimento_Cupom_Fiscal, Funcionario"
''        lSQL = lSQL & "  WHERE Movimento_Cupom_Fiscal.Empresa = " & g_empresa
''        lSQL = lSQL & "    AND Movimento_Cupom_Fiscal.[Codigo do Produto] = " & rstVendaCupomFiscal![Codigo do Produto]
''
''        If chkDetalhadaData.Value = 1 Then
''           ' lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.Data >= " & preparaData(lDataDetalhada)
''            lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.Data = " & preparaData(lDataDetalhada)
''        Else
''            lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.Data >= " & preparaData(msk_data_i.Text)
''            lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.Data <= " & preparaData(msk_data_f.Text)
''        End If
''
'''        lSQL = lSQL & "    AND Movimento_Cupom_Fiscal.Data >= " & preparaData(msk_data_i.Text)
'''        lSQL = lSQL & "    AND Movimento_Cupom_Fiscal.Data <= " & preparaData(msk_data_f.Text)
''        lSQL = lSQL & "    AND Movimento_Cupom_Fiscal.Periodo >= " & preparaTexto(cbo_periodo_i.Text)
''        lSQL = lSQL & "    AND Movimento_Cupom_Fiscal.Periodo <= " & preparaTexto(cbo_periodo_f.Text)
'''        lSQL = lSQL & "    AND ( Movimento_Cupom_Fiscal.[Cupom Cancelado] = " & preparaBooleano(True)
'''        lSQL = lSQL & "    OR Movimento_Cupom_Fiscal.[Item Cancelado] = " & preparaBooleano(True) & ")"
''        lSQL = lSQL & "    AND Movimento_Cupom_Fiscal.[Cupom Cancelado] = " & preparaBooleano(False)
''        lSQL = lSQL & "    AND Movimento_Cupom_Fiscal.[Item Cancelado] = " & preparaBooleano(False)
''
''        If cboFuncionario.ListIndex > 0 Then
''            lSQL = lSQL & "    AND Movimento_Cupom_Fiscal.Operador = " & Val(cboFuncionario.ItemData(cboFuncionario.ListIndex))
''        End If
''        If cboFormaPagamento.ListIndex > 0 Then
''            lSQL = lSQL & "    AND [Forma de Pagamento] = " & Val(cboFormaPagamento.ItemData(cboFormaPagamento.ListIndex))
''        End If
''        If cboProduto.ListIndex > 0 Then
''            lSQL = lSQL & " AND [Codigo do Produto] = " & Val(cboProduto.ItemData(cboProduto.ListIndex))
''        End If
''        If cboCliente.ListIndex > 0 Then
''            lSQL = lSQL & " AND [Codigo do Cliente] = " & Val(cboCliente.ItemData(cboCliente.ListIndex))
''        End If
''        If cboGrupoProduto.ListIndex > 0 Then
''            lSQL = lSQL & " AND Movimento_Cupom_Fiscal.[Codigo do Grupo] = " & Val(cboGrupoProduto.ItemData(cboGrupoProduto.ListIndex))
''        End If
''        If cboTributacao.ListIndex > 0 Then
''            lSQL = lSQL & " AND Movimento_Cupom_Fiscal.[Codigo da Aliquota] = " & Val(cboGrupoProduto.ItemData(cboTributacao.ListIndex))
''        End If
''        If chkSomenteCancelado.Value = 1 Then
''            lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.[Cupom Cancelado] = " & preparaBooleano(True)
''            lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.[Item Cancelado] = " & preparaBooleano(True)
''        Else
''            lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.[Cupom Cancelado] = " & preparaBooleano(False)
''            lSQL = lSQL & "   AND Movimento_Cupom_Fiscal.[Item Cancelado] = " & preparaBooleano(False)
''        End If
''        lSQL = lSQL & "    AND Produto.Codigo = Movimento_Cupom_Fiscal.[Codigo do Produto]"
''        lSQL = lSQL & "    AND Funcionario.Empresa = " & g_empresa
''        lSQL = lSQL & "    AND Funcionario.Codigo = Movimento_Cupom_Fiscal.Operador"
''
''        lSQL = lSQL & "  ORDER BY Movimento_Cupom_Fiscal.Data, Movimento_Cupom_Fiscal.Periodo, Movimento_Cupom_Fiscal.Operador"
''        Set rs = Conectar.RsConexao(lSQL)
''        If rs.RecordCount > 0 Then
''            Do Until rs.EOF
''                ImpDet2
''                lSubQuantidade = rs!Quantidade
''                lSubValor = rs![Valor Total] - rs![Valor do Desconto]
''
''                lSubTotalQuantidade = lSubTotalQuantidade + lSubQuantidade
''                lSubTotalValor = lSubTotalValor + lSubValor
''                lTotalQuantidade = lTotalQuantidade + lSubQuantidade
''                lTotalValor = lTotalValor + lSubValor
''                rs.MoveNext
''            Loop
''        End If
''        rs.Close
''        If lSubValor > 0 Then
''            lSubQuantidade = 0
''            lSubValor = 0
''        End If
''    Else

        'lSubQuantidade = MovimentoCupomFiscal.QuantidadeProdutoVendaData(g_empresa, rstVendaCupomFiscal!Codigo, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), Val(cboFormaPagamento.ItemData(cboFormaPagamento.ListIndex)))
        lSubQuantidade = rstVendaCupomFiscal!TotQuantidade
        'lSubValor = MovimentoCupomFiscal.ValorProdutoVendaData(g_empresa, rstVendaCupomFiscal!Codigo, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), Val(cboFormaPagamento.ItemData(cboFormaPagamento.ListIndex)))
        'lSubValor = lSubValor - MovimentoCupomFiscal.DescontoProdutoVendaData(g_empresa, rstVendaCupomFiscal!Codigo, CDate(msk_data_i.Text), CDate(msk_data_f.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text), Val(cboFormaPagamento.ItemData(cboFormaPagamento.ListIndex)))
        lSubValor = rstVendaCupomFiscal!TotTotal
        
        lSubTotalQuantidade = lSubTotalQuantidade + lSubQuantidade
        lSubTotalValor = lSubTotalValor + lSubValor
        lTotalQuantidade = lTotalQuantidade + lSubQuantidade
        lTotalValor = lTotalValor + lSubValor
        
        If lSubValor <> 0 Then
            ImpDet
            lSubQuantidade = 0
            lSubValor = 0
        End If
''    End If
End Sub
Private Sub ImpDet()
    Dim x_linha As String
    Dim i As Integer
    Dim x_valor As Currency
    
    If lPagina = 0 Then
        ImpCab
    End If
    If lImprimirECF = False Then
        If lLinha >= 60 Then
            x_linha = "+-----------+----------------------------------------------+---------+-----------------+--------------------+--------------------+------+"
            Mid(x_linha, 15, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & x_linha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        x_valor = lSubValor / lSubQuantidade
        x_linha = "|           |                                              |         |                 |                    |                    |      |"
        i = Len(Format(rstVendaCupomFiscal![Codigo do Produto], "#,000"))
        Mid(x_linha, 5 + 5 - i, i) = Format(rstVendaCupomFiscal![Codigo do Produto], "#,000")
        Mid(x_linha, 17, 40) = rstVendaCupomFiscal!NomeProduto
        Mid(x_linha, vbInformation, 3) = rstVendaCupomFiscal!Unidade
        i = Len(Format(lSubQuantidade, "####,##0.00"))
        Mid(x_linha, 74 + 11 - i, i) = Format(lSubQuantidade, "####,##0.00")
        i = Len(Format(x_valor, "###,###,##0.00"))
        Mid(x_linha, 92 + 14 - i, i) = Format(x_valor, "###,###,##0.00")
        i = Len(Format(lSubValor, "###,###,##0.00"))
        Mid(x_linha, 113 + 14 - i, i) = Format(lSubValor, "###,###,##0.00")
        If Aliquota.LocalizarCodigo(lSerieECF, rstVendaCupomFiscal![Codigo da Aliquota]) Then
            If Aliquota.Aliquota > 0 Then
                i = Len(Format(Aliquota.Aliquota, "#0.00"))
                Mid(x_linha, 131 + 5 - i, i) = Format(Aliquota.Aliquota, "###,###,##0.00")
                Mid(x_linha, 136, 1) = "%"
            Else
                If Aliquota.CodigoFiscal = "NN" Then
                    Mid(x_linha, 131, 6) = "N.Inci"
                ElseIf Aliquota.CodigoFiscal = "FF" Then
                    Mid(x_linha, 131, 6) = "S.Trib"
                ElseIf Aliquota.CodigoFiscal = "II" Then
                    Mid(x_linha, 131, 6) = "Isenc."
                Else
                    Mid(x_linha, 131, 2) = Aliquota.CodigoFiscal
                End If
            End If
        End If
        BioImprime "@@Printer.FontName = Courier New"
        BioImprime "@@Printer.FontSize = 7"
        BioImprime "@Printer.Print " & x_linha
    Else
        '          123456789012345678901234567890123456789012345678
        x_linha = "123 1234567890123456789012345678901 152 11487,33"
        x_linha = "                                                "
        Mid(x_linha, 1, 3) = Format(rstVendaCupomFiscal![Codigo do Produto], "000")
        Mid(x_linha, 5, 31) = rstVendaCupomFiscal!NomeProduto
        i = Len(Format(lSubQuantidade, "##0"))
        Mid(x_linha, 37 + 3 - i, i) = Format(lSubQuantidade, "##0")
        i = Len(Format(lSubValor, "####0.00"))
        Mid(x_linha, 41 + 8 - i, i) = Format(lSubValor, "####0.00")
        
        If fUsaNFCe Then 'alex - termica
            Call ImpTermicaImprimeDados(x_linha, True)
        
        ElseIf lImpBematech Then
            BemaRetorno = Bematech_FI_UsaRelatorioGerencialMFD(x_linha)
        ElseIf lImpQuick Then
            'Imprime detalhes do relatorio gerencial
            If EcfQuickImprimeTexto(x_linha) Then
                BemaRetorno = 1
            Else
                BemaRetorno = -1
            End If
        ElseIf lImpDaruma Then
            Call CriaLogCupom("ImpDet - Daruma_FI_RelatorioGerencial. x_linha=" & x_linha)
            BemaRetorno = Daruma_FI_RelatorioGerencial(x_linha)
            Call CriaLogCupom("ImpDet - Daruma_FI_RelatorioGerencial. BemaRetorno=" & BemaRetorno)
        End If
    End If
    lLinha = lLinha + 1
End Sub
Private Sub ImpTermicaAbreRelatorio()
    lNomeArquivo = BioCriaImprime
    'seleciona medidas para centímetros
    BioImprime "@@Printer.ScaleMode = 7"
    BioImprime "@@Printer.PaperSize = 1"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    'teste para imprimir letra correta
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    BioImprime "@@Printer.FontName = Draft 5cpi"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.CurrentY = 0"
    BioImprime "@@Printer.Print " & "  "
    Printer.FontName = "Sans Serif 10cpi"
    Printer.FontName = "Lucida Console 7cpi"
    BioImprime "@@Printer.FontName = Lucida Console 7cpi"
    BioImprime "@@Printer.CurrentY = 0"
End Sub
Private Sub ImpTermicaImprimeDados(ByVal pLinhaDados As String, ByVal pNegrito As Boolean)
    Dim xNegrito As String
    
    If pNegrito = True Then
        xNegrito = "True"
    Else
        xNegrito = "False"
    End If
    BioImprime "@Printer.Print " & pLinhaDados
    BioImprime "@@Printer.FontBold = " & xNegrito
End Sub
Private Sub ImpTermicaFechaRelatorio(ByVal pNomeRelatorio)
    lLocalImpressaoTermica = 1
    BioImprime "@Printer.Print  "
    BioImprime "@Printer.Print  "
    BioImprime "@@Printer.EndDoc"
    BioFechaImprime
    g_string = lLocalImpressaoTermica & lNomeArquivo & "|@|" & pNomeRelatorio & "|@|"
    frm_preview.Show 1
End Sub
Private Sub ImpTermicaRodape()
    Call ImpTermicaImprimeDados("------------------------------------------------", True)
    Call ImpTermicaImprimeDados("Cerrado Tecnologia", True)
    Call ImpTermicaImprimeDados(" ", True)
End Sub
Private Sub ImpTermicaCabecalhoPosto()
    Dim xEmpresa As New cEmpresa
    Dim xLinhaDados As String
    Dim xDados As String
    Dim i As Integer
    
    If xEmpresa.LocalizarCodigo(g_empresa) = False Then
        Exit Sub
    End If
    Call ImpTermicaImprimeDados("------------------------------------------------", True)
    Call ImpTermicaImprimeDados(" ", True)
    
    xLinhaDados = Space(48)
    i = Len(Trim(xEmpresa.Nome))
    Mid(xLinhaDados, 4 + ((40 - i) / 2), i) = Trim(xEmpresa.Nome)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    xLinhaDados = Space(48)
    xDados = "CNPJ: " & fMascaraCNPJ(xEmpresa.CGC) & "  IE: " & xEmpresa.InscricaoEstadual
    i = Len(Trim(xDados))
    Mid(xLinhaDados, 4 + ((40 - i) / 2), i) = Trim(xDados)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    xLinhaDados = Space(48)
    xDados = xEmpresa.Endereco
    i = Len(Trim(xDados))
    Mid(xLinhaDados, 4 + ((40 - i) / 2), i) = Trim(xDados)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    xLinhaDados = Space(48)
    xDados = xEmpresa.Cidade & "-" & xEmpresa.Estado
    i = Len(Trim(xDados))
    Mid(xLinhaDados, 4 + ((40 - i) / 2), i) = Trim(xDados)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    xLinhaDados = Space(48)
    xDados = "CEP: " & fMascaraCEP(xEmpresa.CEP) & "  FONE: " & fMascaraTelefone(xEmpresa.Telefone)
    i = Len(Trim(xDados))
    Mid(xLinhaDados, 4 + ((40 - i) / 2), i) = Trim(xDados)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    Call ImpTermicaImprimeDados(" ", True)
    Call ImpTermicaImprimeDados("------------------------------------------------", True)
    xLinhaDados = "      V E N D A    D E    P R O D U T O S       "
    
    Call ImpTermicaImprimeDados(" ", True)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    Call ImpTermicaImprimeDados("DATA: " & Format(Date, "dd/MM/yyyy") & " AS " & Format(Now, "HH:mm:ss"), True)
    
''    Dim xNomeFuncionario As String
''    xNomeFuncionario = cboFuncionario.Text
''    xLinhaDados = "Funcionario: 000                                "
''    Mid(xLinhaDados, 14, 3) = Format(cboFuncionario.ItemData(cboFuncionario.ListIndex), "000")
''    Mid(xLinhaDados, 18, Len(xNomeFuncionario)) = xNomeFuncionario
    
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    Call ImpTermicaImprimeDados(" ", True)
    Call ImpTermicaImprimeDados("------------------------------------------------", True)
End Sub


Private Sub ImpDet2()
    Dim x_linha As String
    Dim i As Integer
    Dim x_valor As Currency
    
   
    If lPagina = 0 Then
        ImpCab
    End If
    If lLinha >= 60 Then
        x_linha = "+-----+----------------------------------------+---+-----------+--------------+--------------+--------+---+--+----------+---------------+"
        Mid(x_linha, 15, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    '                   1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '          12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    x_linha = "|     |                                        |   |           |              |              |        |   |  |          |               |"
    i = Len(Format(rstVendaCupomFiscal![Codigo do Produto], "#,000"))
    Mid(x_linha, 2 + 5 - i, i) = Format(rstVendaCupomFiscal![Codigo do Produto], "#,000")
    Mid(x_linha, 8, 40) = rs!Nome
    Mid(x_linha, 49, 3) = rs!Unidade
    i = Len(Format(rs!Quantidade, "####,##0.00"))
    Mid(x_linha, 53 + 11 - i, i) = Format(rs!Quantidade, "####,##0.00")
    i = Len(Format(rs![Valor Unitario], "###,###,##0.00"))
    Mid(x_linha, 65 + 14 - i, i) = Format(rs![Valor Unitario], "###,###,##0.00")
    x_valor = rs![Valor Total] - rs![Valor do Desconto]
    i = Len(Format(x_valor, "###,###,##0.00"))
    Mid(x_linha, 80 + 14 - i, i) = Format(x_valor, "###,###,##0.00")
    i = Len(Format(rs![Numero do Cupom], "####,##0"))
    Mid(x_linha, 95 + 8 - i, i) = Format(rs![Numero do Cupom], "####,##0")
    i = Len(Format(rs!Ordem, "##0"))
    Mid(x_linha, 104 + 3 - i, i) = Format(rs!Ordem, "##0")
    i = Len(Format(rs![Tipo do SubEstoque], "#0"))
    Mid(x_linha, 107 + 3 - i, i) = Format(rs![Tipo do SubEstoque], "#0")
    Mid(x_linha, 111, 10) = Format(rs!Data, "dd/mm/yyyy")
    Mid(x_linha, 122, 15) = rs!NomeFuncionario
    BioImprime "@@Printer.FontName = Courier New"
    BioImprime "@@Printer.FontSize = 7"
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpSubTotal()
    Dim x_linha As String
    Dim i As Integer
    If lImprimirECF = False Then
''        If chkImprimeDetalhe.Value = 1 Then
''            x_linha = "+-----+----------------------------------------+---+-----------+--------------+--------------+--------+---+--+----------+---------------+"
''            BioImprime "@Printer.Print " & x_linha
''            x_linha = "|  *** TOTAL DO GRUPO:                             |           |              |              |                                          |"
''            Mid(x_linha, 24, 28) = lNomeGrupo
''            i = Len(Format(lSubTotalQuantidade, "####,##0.00"))
''            Mid(x_linha, 53 + 11 - i, i) = Format(lSubTotalQuantidade, "####,##0.00")
''            i = Len(Format(lSubTotalValor, "###,###,##0.00"))
''            Mid(x_linha, 80 + 14 - i, i) = Format(lSubTotalValor, "###,###,##0.00")
''            BioImprime "@Printer.Print " & x_linha
''            x_linha = "+-----+----------------------------------------+---+-----------+--------------+--------------+--------+---+--+----------+---------------+"
''            BioImprime "@Printer.Print " & x_linha
''        Else
            x_linha = "+-----------+----------------------------------------------+---------+-----------------+--------------------+--------------------+------+"
            BioImprime "@Printer.Print " & x_linha
            x_linha = "|                 *** TOTAL DO GRUPO:                                |                 |                    |                    |      |"
            Mid(x_linha, 39, 30) = lNomeGrupo
            i = Len(Format(lSubTotalQuantidade, "####,##0.00"))
            Mid(x_linha, 74 + 11 - i, i) = Format(lSubTotalQuantidade, "####,##0.00")
            i = Len(Format(lSubTotalValor, "###,###,##0.00"))
            Mid(x_linha, 113 + 14 - i, i) = Format(lSubTotalValor, "###,###,##0.00")
            BioImprime "@Printer.Print " & x_linha
            x_linha = "+-----------+----------------------------------------------+---------+-----------------+--------------------+--------------------+------+"
            BioImprime "@Printer.Print " & x_linha
''        End If
    End If
End Sub
Private Sub ImpTotal()
'    Dim y_local As Single
    Dim x_linha As String
    Dim i As Integer
    If lImprimirECF = False Then
''        If chkImprimeDetalhe.Value = 1 Then
''
''            If chkDetalhadaData.Value = 1 Then
''                x_linha = "|               *** TOTAL GERAL DA DATA            |           |              |              |                                          |"
''                Mid(x_linha, 41, 10) = lDataDetalhada
''            Else
''                x_linha = "|                 *** TOTAL GERAL                  |           |              |              |                                          |"
''            End If
''
''               'x_linha = "|                                                  |           |              |              |                                          |"
''            i = Len(Format(lTotalQuantidade, "####,##0.00"))
''            Mid(x_linha, 53 + 11 - i, i) = Format(lTotalQuantidade, "####,##0.00")
''            i = Len(Format(lTotalValor, "###,###,##0.00"))
''            Mid(x_linha, 80 + 14 - i, i) = Format(lTotalValor, "###,###,##0.00")
''            BioImprime "@@Printer.FontName = Courier New"
''            BioImprime "@@Printer.FontSize = 7"
''            BioImprime "@@y_local = Printer.CurrentY"
''            BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
''            BioImprime "@@Printer.CurrentY = y_local"
''            BioImprime "@@Printer.FontBold = True"
''            BioImprime "@Printer.Print " & x_linha
''        '    Printer.CurrentY = y_local - 0.01
''        '    Printer.Print x_linha
''            BioImprime "@@Printer.CurrentY = y_local"
''            BioImprime "@@Printer.Print " & Chr(34) & " " & Chr(34)
''            BioImprime "@@Printer.FontBold = False"
''            x_linha = "+--------------------------------------------------+-----------+--------------+--------------+------------------------------------------+"
''            'Mid(x_linha, 3, 22) = " Cerrado Informática. "
''            BioImprime "@Printer.Print " & x_linha
''            BioImprime "@@Printer.FontName = Courier New"
''            BioImprime "@Printer.Print " & " "
''        Else
        
''            If chkDetalhadaData.Value = 1 Then
''                x_linha = "|                 *** TOTAL GERAL DA DATA                            |                 |                    |                    |      |"
''                Mid(x_linha, 43, 10) = lDataDetalhada
''            Else
                x_linha = "|                 *** TOTAL GERAL                                    |                 |                    |                    |      |"
''            End If
            
            'x_linha = "|                 *** TOTAL GERAL                                    |                 |                    |                    |      |"
            i = Len(Format(lTotalQuantidade, "####,##0.00"))
            Mid(x_linha, 74 + 11 - i, i) = Format(lTotalQuantidade, "####,##0.00")
            i = Len(Format(lTotalValor, "###,###,##0.00"))
            Mid(x_linha, 113 + 14 - i, i) = Format(lTotalValor, "###,###,##0.00")
            BioImprime "@@Printer.FontName = Courier New"
            BioImprime "@@Printer.FontSize = 7"
            BioImprime "@@y_local = Printer.CurrentY"
            BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
            BioImprime "@@Printer.CurrentY = y_local"
            BioImprime "@@Printer.FontBold = True"
            BioImprime "@Printer.Print " & x_linha
        '    Printer.CurrentY = y_local - 0.01
        '    Printer.Print x_linha
            BioImprime "@@Printer.CurrentY = y_local"
            BioImprime "@@Printer.Print " & Chr(34) & " " & Chr(34)
            BioImprime "@@Printer.FontBold = False"
            x_linha = "+--------------------------------------------------------------------+-----------------+--------------------+--------------------+------+"
            'Mid(x_linha, 3, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & x_linha
''        End If
    Else
        '          123456789012345678901234567890123456789012345678
        x_linha = "123 1234567890123456789012345678901 152 11487,33"
        
''        If chkDetalhadaData.Value = 1 Then
''            x_linha = "|                 *** TOTAL GERAL DA DATA                            |                 |                    |                    |      |"
''            Mid(x_linha, 43, 10) = lDataDetalhada
''        Else
            x_linha = "                  *** TOTAL GERAL                                    |                 |                    |                    |      |"
''        End If
        
        'x_linha = "                    *** TOTAL GERAL             "
        i = Len(Format(lTotalQuantidade, "##0"))
        Mid(x_linha, 37 + 3 - i, i) = Format(lTotalQuantidade, "##0")
        i = Len(Format(lTotalValor, "####0.00"))
        Mid(x_linha, 41 + 8 - i, i) = Format(lTotalValor, "####0.00")
        
        If fUsaNFCe Then 'alex - termica
            Call ImpTermicaImprimeDados(x_linha, True)
        ElseIf lImpBematech Then
            BemaRetorno = Bematech_FI_UsaRelatorioGerencialMFD(x_linha)
        ElseIf lImpQuick Then
            'Imprime detalhes do relatorio gerencial
            If EcfQuickImprimeTexto(x_linha) Then
                BemaRetorno = 1
            Else
                BemaRetorno = -1
            End If
        ElseIf lImpDaruma Then
            Call CriaLogCupom("ImpTotal - Daruma_FI_AbreRelatorioGerencial.")
            BemaRetorno = Daruma_FI_AbreRelatorioGerencial()
            Call CriaLogCupom("ImpTotal - Daruma_FI_AbreRelatorioGerencial. BemaRetorno=" & BemaRetorno)
        End If
    End If
End Sub
Private Sub ImpCab()
    Dim x_linha As String
    Dim i As Integer
    If lImprimirECF = False Then
        If lPagina = 0 Then
            lNomeArquivo = BioCriaImprime
            'seleciona medidas para centímetros
            BioImprime "@@Printer.ScaleMode = 7"
            BioImprime "@@Printer.PaperSize = 1"
            BioImprime "@@Printer.FontName = Courier New"
            BioImprime "@@Printer.FontName = Courier New"
            'teste para imprimir letra correta
            BioImprime "@@Printer.FontBold = False"
            BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
        End If
        lPagina = lPagina + 1
        lLinha = 0
        BioImprime "@@Printer.FontName = Draft 5cpi"
        BioImprime "@@Printer.FontName = Courier New"
        BioImprime "@@Printer.CurrentY = 0"
        BioImprime "@@Printer.Print " & Chr(34) & " " & Chr(34)
        BioImprime "@@Printer.FontName = Courier New"
        BioImprime "@@Printer.FontSize = 10"
        BioImprime "@@Printer.CurrentY = 0"
        '                   1         2         3         4         5         6         7         8
        '          12345678901234567890123456789012345678901234567890123456789012345678901234567890
        x_linha = "+------------------------------------------------------------------------------+"
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.FontBold = True"
        x_linha = "|                                                                  Página: ___ |"
        Mid(x_linha, 3, 40) = g_nome_empresa
        Mid(x_linha, 76, 3) = Format(lPagina, "000")
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.FontBold = False"
        x_linha = "| VENDAS DO CUPOM FISCAL                                          , __/__/____ |"
        i = Len(g_cidade_empresa)
        Mid(x_linha, 37 + 30 - i, i) = g_cidade_empresa
        Mid(x_linha, 69, 10) = msk_data.Text
        BioImprime "@Printer.Print " & x_linha
        x_linha = "| REFERENTE AO PERIODO DE.: __/__/____ A __/__/____.                           |"
        Mid(x_linha, 29, 10) = msk_data_i.Text
        Mid(x_linha, 42, 10) = msk_data_f.Text
        BioImprime "@Printer.Print " & x_linha
        x_linha = "| CAIXA INICIAL...........: X    CAIXA FINAL..: X                              |"
        Mid(x_linha, 29, 1) = cbo_periodo_i.Text
        Mid(x_linha, 49, 1) = cbo_periodo_f.Text
        BioImprime "@Printer.Print " & x_linha
        x_linha = "| FUNCIONARIO.............:                                                    |"
        Mid(x_linha, 29, 30) = "Todos Funcionários"
        BioImprime "@Printer.Print " & x_linha
''        x_linha = "| FORMA DE PAGAMENTO......:                                                    |"
''        Mid(x_linha, 29, 30) = cboFormaPagamento.Text
''        BioImprime "@Printer.Print " & x_linha
''        BioImprime "@@Printer.FontName = Courier New"
''        BioImprime "@@Printer.FontSize = 7"
''        If chkImprimeDetalhe.Value = 1 Then
''            x_linha = "+-----+----------------------------------------+---+-----------+--------------+--------------+--------+---+--+----------+---------------+"
''            BioImprime "@Printer.Print " & x_linha
''            x_linha = "|CODGO|DISCRIMINAÇÃO DOS PRODUTOS              |UND| QUANTIDADE|PRECO DE VENDA|TOTAL DA VENDA|N. CUPOM|ORD|SE|   DATA   |NOME DO FUNC.  |"
''            BioImprime "@Printer.Print " & x_linha
''            x_linha = "+-----+----------------------------------------+---+-----------+--------------+--------------+--------+---+--+----------+---------------+"
''            BioImprime "@Printer.Print " & x_linha
''        Else
            x_linha = "+-----------+----------------------------------------------+---------+-----------------+--------------------+--------------------+------+"
            BioImprime "@Printer.Print " & x_linha
            x_linha = "|  CODIGO   |   DISCRIMINAÇÃO DOS PRODUTOS                 | UNIDADE |  QTD.   VENDAS  |  PRECO  DE  VENDA  |  TOTAL  DA  VENDA  |      |"
            BioImprime "@Printer.Print " & x_linha
            x_linha = "+-----------+----------------------------------------------+---------+-----------------+--------------------+--------------------+------+"
            BioImprime "@Printer.Print " & x_linha
''        End If
    Else
        If lPagina = 0 Then
        '              123456789012345678901234567890123456789012345678
            x_linha = "      V E N D A    D E    P R O D U T O S       "
            If fUsaNFCe Then 'alex - termica
             Call ImpTermicaAbreRelatorio
             Call ImpTermicaCabecalhoPosto
            ElseIf lImpBematech Then
                BemaRetorno = Bematech_FI_AbreRelatorioGerencialMFD("01")
                BemaRetorno = Bematech_FI_UsaRelatorioGerencialMFD(x_linha)
            ElseIf lImpQuick Then
                'Abre Relatorio Gerencial DATAREGIS
                If EcfQuickDefineGerencial(0, "Gerencial") Then
                    BemaRetorno = 1
                    If EcfQuickAbreGerencial(0, "Gerencial") Then
                        BemaRetorno = 1
                    Else
                        BemaRetorno = 0
                    End If
                Else
                    BemaRetorno = 0
                End If
                'Imprime detalhes do relatorio gerencial
                If EcfQuickImprimeTexto(x_linha) Then
                    BemaRetorno = 1
                Else
                    BemaRetorno = -1
                End If
            ElseIf lImpDaruma Then
                Call CriaLogCupom("ImpCab - Daruma_FI_RelatorioGerencial. x_linha=" & x_linha)
                BemaRetorno = Daruma_FI_RelatorioGerencial(x_linha)
                Call CriaLogCupom("ImpCab - Daruma_FI_RelatorioGerencial. BemaRetorno=" & BemaRetorno)
            End If
            x_linha = "      REFERENTE A: __/__/____ A __/__/____      "
            Mid(x_linha, 20, 10) = msk_data_i.Text
            Mid(x_linha, 33, 10) = msk_data_f.Text
            
            If fUsaNFCe Then 'ALEX - TERMICA
                Call ImpTermicaImprimeDados(x_linha, True)
            ElseIf lImpBematech Then
                BemaRetorno = Bematech_FI_UsaRelatorioGerencialMFD(x_linha)
            ElseIf lImpQuick Then
                'Imprime detalhes do relatorio gerencial
                If EcfQuickImprimeTexto(x_linha) Then
                    BemaRetorno = 1
                Else
                    BemaRetorno = -1
                End If
            ElseIf lImpDaruma Then
                Call CriaLogCupom("ImpCab - Daruma_FI_RelatorioGerencial. x_linha=" & x_linha)
                BemaRetorno = Daruma_FI_RelatorioGerencial(x_linha)
                Call CriaLogCupom("ImpCab - Daruma_FI_RelatorioGerencial. BemaRetorno=" & BemaRetorno)
            End If
            x_linha = "      CAIXA: X AO X                             "
            Mid(x_linha, 14, 1) = cbo_periodo_i.Text
            Mid(x_linha, 19, 1) = cbo_periodo_f.Text
            
            If fUsaNFCe Then 'alex - termica
                Call ImpTermicaImprimeDados(x_linha, True)
            ElseIf lImpBematech Then
                BemaRetorno = Bematech_FI_UsaRelatorioGerencialMFD(x_linha)
            ElseIf lImpQuick Then
                'Imprime detalhes do relatorio gerencial
                If EcfQuickImprimeTexto(x_linha) Then
                    BemaRetorno = 1
                Else
                    BemaRetorno = -1
                End If
            ElseIf lImpDaruma Then
                Call CriaLogCupom("ImpCab - Daruma_FI_RelatorioGerencial. x_linha=" & x_linha)
                BemaRetorno = Daruma_FI_RelatorioGerencial(x_linha)
                Call CriaLogCupom("ImpCab - Daruma_FI_RelatorioGerencial. BemaRetorno=" & BemaRetorno)
            End If
            x_linha = "COD NOME PRODUTO                    QTD VL.TOTAL"
            If fUsaNFCe Then
                Call ImpTermicaImprimeDados(x_linha, True)
            ElseIf lImpBematech Then
                BemaRetorno = Bematech_FI_UsaRelatorioGerencialMFD(x_linha)
            ElseIf lImpQuick Then
                'Imprime detalhes do relatorio gerencial
                If EcfQuickImprimeTexto(x_linha) Then
                    BemaRetorno = 1
                Else
                    BemaRetorno = -1
                End If
            ElseIf lImpDaruma Then
                Call CriaLogCupom("ImpCab - Daruma_FI_RelatorioGerencial. x_linha=" & x_linha)
                BemaRetorno = Daruma_FI_RelatorioGerencial(x_linha)
                Call CriaLogCupom("ImpCab - Daruma_FI_RelatorioGerencial. BemaRetorno=" & BemaRetorno)
            End If
        End If
        lPagina = lPagina + 1
    End If
End Sub
Private Sub cbo_periodo_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboTipoVenda.SetFocus
    End If
End Sub
Private Sub cbo_periodo_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If lUtilizaNFCe Then
            cboTipoVenda.SetFocus
        Else
            cbo_periodo_f.ListIndex = cbo_periodo_i.ListIndex
            cbo_periodo_f.SetFocus
        End If
    End If
End Sub
Private Sub cboFormaPagamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
''Private Sub cboFuncionario_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        KeyAscii = 0
''        cboFormaPagamento.SetFocus
''    End If
''End Sub
Private Sub cboTipoVenda_KeyPress(KeyAscii As Integer)
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
        cmd_visualizar.SetFocus
    Else
        msk_data.Text = RetiraGString(1)
        msk_data_i.SetFocus
    End If
    g_string = " "
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
    g_string = " "
    cmd_visualizar.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cmd_visualizar.SetFocus
    Else
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    lLocal = 1
    lImprimirECF = False
    
    If ValidaCampos Then
        If lEcfInstalada Then
            If fUsaNFCe Then 'alex - termica
                
''                If chkDetalhadaData.Value = 1 Or chkImprimeDetalhe.Value = 1 Then
''                    MsgBox "Este relatório não pode ser impresso de forma detalhada na impressora térmica", vbInformation + vbOKOnly + vbDefaultButton1, "Operação não permitida"
''                    Exit Sub
''                End If
                
                If (MsgBox("Deseja imprimir na impressora térmica?", vbQuestion + vbYesNo + vbDefaultButton2, "Forma de Impressão!")) = vbYes Then
                    lImprimirECF = True
                End If
            Else
                If (MsgBox("Deseja imprimir na impressora fiscal?", vbQuestion + vbYesNo + vbDefaultButton2, "Forma de Impressão!")) = vbYes Then
                    lImprimirECF = True
                End If
            End If
        
        End If
        If lEcfInstalada And lImprimirECF Then
            Call GravaAuditoria(1, Me.name, 7, IIf(fUsaNFCe, "Para Térmica", "Para ECF"))
            If fUsaNFCe Then
                'Call SelecionaImpressoraPadrao("Gerando Relatório!", Me.Caption, Me.Top, Me.Left, Me.Width, Me.Height)
                DefineImpressoraTermicaComoPadrao
            End If
            
            LoopRelatorio
        Else
            If SelecionaImpressoraHP(Me) Then
                Call GravaAuditoria(1, Me.name, 7, "")
                LoopRelatorio
            End If
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
    ElseIf cbo_periodo_i.ListIndex = -1 Then
        MsgBox "Selecione o período inicial.", vbInformation, "Atenção!"
        cbo_periodo_i.SetFocus
    ElseIf cbo_periodo_f.ListIndex = -1 Then
        MsgBox "Selecione o período final.", vbInformation, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cbo_periodo_f.Text < cbo_periodo_i.Text Then
        MsgBox "Periodo final deve ser maior.", vbInformation, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cboTipoVenda.ListIndex = -1 Then
        MsgBox "Selecione o tipo de venda.", vbInformation, "Atenção!"
        cboTipoVenda.SetFocus
''    ElseIf cboFormaPagamento.ListIndex = -1 Then
''        MsgBox "Selecione uma forma de pagamento.", vbInformation, "Atenção!"
''        cboFormaPagamento.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    lLocal = 0
    lImprimirECF = False
    If ValidaCampos Then
        If SelecionaImpressoraHP(Me) Then
            Call GravaAuditoria(1, Me.name, 6, "")
            
            LoopRelatorio
        End If
    End If
End Sub
Private Sub cmd_visualizar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 18 Then 'Crtl + R
        KeyAscii = 0
        ZzLoopRecalculaDescontoCupom
    End If
End Sub

Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def, "dd/mm/yyyy")
        cbo_periodo_i.ListIndex = 0
        cbo_periodo_f.ListIndex = 3
        cboTipoVenda.ListIndex = 0
''        cboFuncionario.ListIndex = 0
''        cboFormaPagamento.ListIndex = 0
''        cboProduto.ListIndex = 0
''        cboCliente.ListIndex = 0
''        cboGrupoProduto.ListIndex = 0
''        cboTributacao.ListIndex = 0
        If UCase(g_nome_usuario) Like "*LOJA*" Then
            cboTipoVenda.ListIndex = 1
        End If
        cmd_imprimir.SetFocus
    End If
    Screen.MousePointer = 1
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
    PreencheCboPeriodo
    PreencheCboTipoVenda
''    PreencheCboFormaPagamento
''    PreencheCboFuncionario
''    PreencheCboProduto
''    PreencheCboCliente
''    PreencheCboGrupoProduto
''    PreencheCboTributacao
    AtualizaConstantes
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub



Private Sub msk_data_f_GotFocus()
    msk_data_f.SelStart = 0
    msk_data_f.SelLength = 2
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If lUtilizaNFCe Then
            cboTipoVenda.SetFocus
        Else
            cbo_periodo_i.SetFocus
        End If
    End If
End Sub
Private Sub msk_data_i_GotFocus()
    msk_data_i.SelStart = 0
    msk_data_i.SelLength = 2
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
