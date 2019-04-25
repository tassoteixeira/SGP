VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_movimentacao_estoque 
   Caption         =   "Emissão da Movimentação do Estoque"
   ClientHeight    =   3330
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_movimentacao_estoque.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_movimentacao_estoque.frx":030A
   ScaleHeight     =   3330
   ScaleWidth      =   6795
   Begin VB.Frame frm_dados 
      Height          =   2235
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6555
      Begin VB.CheckBox chkCupomFiscal 
         Caption         =   "NFCe"
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Top             =   1860
         Width           =   3015
      End
      Begin VB.ComboBox cbo_produto 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1500
         Width           =   4755
      End
      Begin VB.ComboBox cbo_grupo 
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
         Picture         =   "lst_movimentacao_estoque.frx":0350
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "lst_movimentacao_estoque.frx":162A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_movimentacao_estoque.frx":2904
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
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
         Left            =   4800
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
      Begin VB.Label Label5 
         Caption         =   "&Data de Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Produto"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Grupo"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Data &final"
         Height          =   315
         Left            =   3840
         TabIndex        =   7
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "D&ata inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_movimentacao_estoque.frx":3BDE
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Visualiza a movimentação do estoque."
      Top             =   2400
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_movimentacao_estoque.frx":52F8
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Imprime a movimentação do estoque."
      Top             =   2400
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_movimentacao_estoque.frx":6902
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2400
      Width           =   795
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_movimentacao_estoque"
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
Dim lSubAnterior As Currency
Dim lSubEntrada As Currency
Dim lSubSaida As Currency
Dim lSubAtual As Currency
Dim lImprimiuCab2 As Boolean
Dim lImprimiuProduto As Boolean
Dim lGrupo As Integer
Dim lData As Date
Dim lSQL As String
Dim lCodigosProdutos As String

Dim rsGrupo As New adodb.Recordset
Dim rsProduto As New adodb.Recordset
Dim rsEntradaProduto As New adodb.Recordset
Dim rsMovimentoLubrificante As New adodb.Recordset

Private Cfop As New cCfop
Private Estoque2 As New cEstoque2
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Cfop = Nothing
    Set Estoque2 = Nothing
End Sub
Private Sub MontaRsEntradaProduto()
    lSQL = ""
    lSQL = lSQL & "   SELECT [Codigo do Produto], [Data da Entrada], [Numero do Documento], CFOP, Quantidade"
    lSQL = lSQL & "     FROM Entrada_Produto"
    lSQL = lSQL & "    WHERE Empresa = " & g_empresa
    lSQL = lSQL & "      AND [Data da Entrada] >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "      AND [Data da Entrada] <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & "      AND [Tipo da Entrada] <> " & 3 'Inventario
    If cbo_produto.ItemData(cbo_produto.ListIndex) <> 0 Then
        lSQL = lSQL & "    AND [Codigo do Produto] = " & Val(cbo_produto.ItemData(cbo_produto.ListIndex))
    End If
    lSQL = lSQL & " ORDER BY [Codigo do Produto], [Data da Entrada], [Numero do Documento], CFOP"
    'Abre RecordSet
    Set rsEntradaProduto = New adodb.Recordset
    Set rsEntradaProduto = Conectar.RsConexao(lSQL)
End Sub
Private Sub MontaRsMovimentoLubrificante()
    lSQL = ""
    
    
    If chkCupomFiscal.Value = 0 Then
        lSQL = lSQL & "   SELECT [Codigo do Produto2], Data, Periodo, Quantidade, [Codigo do Funcionario]"
        lSQL = lSQL & "     FROM Movimento_Lubrificante"
        lSQL = lSQL & "    WHERE Empresa = " & g_empresa
        lSQL = lSQL & "      AND Data >= " & preparaData(msk_data_i.Text)
        lSQL = lSQL & "      AND Data <= " & preparaData(msk_data_f.Text)
        If cbo_produto.ItemData(cbo_produto.ListIndex) <> 0 Then
            lSQL = lSQL & "  AND [Codigo do Produto2] = " & Val(cbo_produto.ItemData(cbo_produto.ListIndex))
        End If
        lSQL = lSQL & " ORDER BY [Codigo do Produto2], Data, Periodo"
    Else
'        lSQL = ""
'        lSQL = lSQL & "   SELECT [Codigo do Produto] As [Codigo do Produto2], Data, Periodo, Quantidade, Operador AS [Codigo do Funcionario]"
'        lSQL = lSQL & "     FROM Movimento_Cupom_Fiscal"
'        lSQL = lSQL & "    WHERE Empresa = " & g_empresa
'        lSQL = lSQL & "      AND Data >= " & preparaData(msk_data_i.Text)
'        lSQL = lSQL & "      AND Data <= " & preparaData(msk_data_f.Text)
'        lSQL = lSQL & "      AND [Item Cancelado] = " & preparaBooleano(False)
'        lSQL = lSQL & "      AND [Cupom Cancelado] = " & preparaBooleano(False)
'        If cbo_produto.ItemData(cbo_produto.ListIndex) <> 0 Then
'            lSQL = lSQL & "  AND [Codigo do Produto] = " & Val(cbo_produto.ItemData(cbo_produto.ListIndex))
'        End If
'        lSQL = lSQL & " ORDER BY [Codigo do Produto], Data, Periodo"
        lSQL = ""
        lSQL = lSQL & "   SELECT IdProduto_MovDEItem As [Codigo do Produto2], DataEmissao_MovDEItem AS Data, 0 AS Periodo, Quantidade_MovDEItem AS Quantidade, IdUsuario_MovDECabecalho AS [Codigo do Funcionario]"
        lSQL = lSQL & "     FROM MovimentoDocumentoEletronicoCabecalho, MovimentoDocumentoEletronicoItem"
        lSQL = lSQL & "    WHERE IdEstabelecimento_MovDEItem = " & g_empresa
        lSQL = lSQL & "      AND DataEmissao_MovDEItem >= " & preparaData(msk_data_i.Text)
        lSQL = lSQL & "      AND DataEmissao_MovDEItem <= " & preparaData(msk_data_f.Text)
        lSQL = lSQL & "      AND Cancelado_MovDEItem = " & preparaBooleano(False)
        lSQL = lSQL & "      AND Cancelado_MovDECabecalho = " & preparaBooleano(False)
        lSQL = lSQL & "      AND IdEstabelecimento_MovDEItem = IdEstabelecimento_MovDECabecalho"
        lSQL = lSQL & "      AND DataEmissao_MovDEItem = DataEmissao_MovDECabecalho"
        lSQL = lSQL & "      AND Numero_MovDEItem = Numero_MovDECabecalho"
        lSQL = lSQL & "      AND Serie_MovDEItem = Serie_MovDECabecalho"
        lSQL = lSQL & "      AND Modelo_MovDEItem = Modelo_MovDECabecalho"
        If cbo_produto.ItemData(cbo_produto.ListIndex) <> 0 Then
            lSQL = lSQL & "  AND IdProduto_MovDEItem = " & Val(cbo_produto.ItemData(cbo_produto.ListIndex))
        End If
        lSQL = lSQL & " ORDER BY IdProduto_MovDEItem, DataEmissao_MovDEItem"
    End If
    'Abre RecordSet
    Set rsMovimentoLubrificante = New adodb.Recordset
    Set rsMovimentoLubrificante = Conectar.RsConexao(lSQL)
End Sub
Private Sub PreencheCboGrupo()
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Codigo, Nome"
    lSQL = lSQL & "     FROM Grupo"
    lSQL = lSQL & " ORDER BY Nome, Codigo"
    'Abre RecordSet
    Set rsGrupo = New adodb.Recordset
    Set rsGrupo = Conectar.RsConexao(lSQL)
    
    cbo_grupo.Clear
    cbo_grupo.AddItem "Todos os Grupos"
    cbo_grupo.ItemData(cbo_grupo.NewIndex) = 0
    If rsGrupo.RecordCount > 0 Then
        rsGrupo.MoveFirst
        Do Until rsGrupo.EOF
            cbo_grupo.AddItem rsGrupo("Nome").Value
            cbo_grupo.ItemData(cbo_grupo.NewIndex) = rsGrupo("Codigo").Value
            rsGrupo.MoveNext
        Loop
    End If
    rsGrupo.Close
    Set rsGrupo = Nothing
End Sub
Private Sub PreencheCboProduto()
    cbo_produto.Clear
    
    cbo_produto.AddItem "Todos os Produtos"
    cbo_produto.ItemData(cbo_produto.NewIndex) = 0
    lSQL = ""
    lSQL = lSQL & "SELECT Codigo, Nome"
    lSQL = lSQL & "  FROM Produto"
    lSQL = lSQL & " ORDER BY Nome, Codigo"
    Set rsProduto = Conectar.RsConexao(lSQL)
    If rsProduto.RecordCount > 0 Then
        Do Until rsProduto.EOF
                cbo_produto.AddItem rsProduto!Nome
                cbo_produto.ItemData(cbo_produto.NewIndex) = rsProduto!Codigo
            rsProduto.MoveNext
        Loop
    End If
    rsProduto.Close
    Set rsProduto = Nothing
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lSubAnterior = 0
    lSubEntrada = 0
    lSubSaida = 0
    lSubAtual = 0
    lGrupo = 0
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'ImpDados
    LoopProdutos
    cmd_sair.SetFocus
End Sub
Private Sub LoopProdutos()
    Dim i As Long
    
    lSQL = ""
    lSQL = lSQL & "   SELECT Produto.Codigo, Produto.Nome, Grupo.Codigo AS CodigoGrupo, Grupo.Nome AS NomeGrupo"
    lSQL = lSQL & "     FROM Produto, Grupo"
    
    lSQL = lSQL & "    WHERE Produto.[Codigo do Grupo] = Grupo.Codigo"
    If cbo_grupo.ItemData(cbo_grupo.ListIndex) <> 0 Then
        lSQL = lSQL & "    AND Produto.[Codigo do Grupo] = " & Val(cbo_grupo.ItemData(cbo_grupo.ListIndex))
    End If
    If cbo_produto.ItemData(cbo_produto.ListIndex) <> 0 Then
        lSQL = lSQL & "    AND Produto.Codigo = " & Val(cbo_produto.ItemData(cbo_produto.ListIndex))
    End If
    
    'Início Select Tabela Entrada e Saida
    lSQL = lSQL & "      AND Produto.Codigo IN"
    lSQL = lSQL & " "
    lSQL = lSQL & " ("
    lSQL = lSQL & "   SELECT [Codigo do Produto] As Codigo2"
    lSQL = lSQL & "     FROM Entrada_Produto"
    lSQL = lSQL & "    WHERE Empresa = " & g_empresa
    lSQL = lSQL & "      AND [Data da Entrada] >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "      AND [Data da Entrada] <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & " GROUP BY [Codigo do Produto]"
    lSQL = lSQL & " "
    lSQL = lSQL & "    UNION"
    lSQL = lSQL & " "
    lSQL = lSQL & "   SELECT [Codigo do Produto2] As Codigo2"
    lSQL = lSQL & "     FROM Movimento_Lubrificante"
    lSQL = lSQL & "    WHERE Empresa = " & g_empresa
    lSQL = lSQL & "      AND Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "      AND Data <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & " GROUP BY [Codigo do Produto2]"
    lSQL = lSQL & " )"
    'Fim Select Tabela Entrada e Saida
    lSQL = lSQL & " ORDER BY Grupo.Nome, Grupo.Codigo, Produto.Nome, Produto.Codigo"
    'Abre RecordSet
    Set rsProduto = New adodb.Recordset
    Set rsProduto = Conectar.RsConexao(lSQL)
    If rsProduto.RecordCount > 0 Then
        lCodigosProdutos = "( "
        i = 0
        rsProduto.MoveFirst
        Do Until rsProduto.EOF
            If i > 0 Then
                lCodigosProdutos = lCodigosProdutos + ", "
            End If
            lCodigosProdutos = lCodigosProdutos & rsProduto("Codigo").Value
            i = i + 1
            rsProduto.MoveNext
        Loop
        lCodigosProdutos = lCodigosProdutos + " )"
    End If
    
    
    If rsProduto.RecordCount > 0 Then
        MontaRsEntradaProduto
        MontaRsMovimentoLubrificante
        rsProduto.MoveFirst
        Do Until rsProduto.EOF
            lSubAnterior = 0
            lSubEntrada = 0
            lSubSaida = 0
            lSubAtual = 0
            
            'If cbo_produto.ItemData(cbo_produto.ListIndex) = 0 Or cbo_produto.ItemData(cbo_produto.ListIndex) = rsProduto("Codigo").Value Then
            If lPagina = 0 Then
                ImpCab
            End If
            
            'Le tabela auxiliar
            If chkCupomFiscal.Value = 1 Then
                If Estoque2.LocalizarCodigo(g_empresa, CDate(msk_data_i.Text) - 1, rsProduto("Codigo").Value) Then
                    lSubAnterior = Estoque2.Quantidade
                    lSubAtual = Estoque2.Quantidade
                End If
            End If
            
            lImprimiuCab2 = False
            lData = CDate(msk_data_i.Text)
            lImprimiuProduto = False
            Do Until lData > CDate(msk_data_f.Text)
                Call LoopRsMovimentoEntrada(rsProduto("Codigo").Value)
                Call LoopRsMovimentoSaida(rsProduto("Codigo").Value)
                lData = lData + 1
            Loop
            If lSubAnterior <> 0 Or lSubEntrada <> 0 Or lSubSaida <> 0 Or lSubAtual Then
                Call ImpSubTotal
            End If
            'End If
            rsProduto.MoveNext
        Loop
    End If
    rsProduto.Close
    Set rsProduto = Nothing
    
    If lPagina > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Movimentação do Estoque|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub BuscaEntrada()

    lSQL = ""
    lSQL = lSQL & "   SELECT [Data da Entrada], [Numero do Documento], CFOP, Quantidade "
    lSQL = lSQL & "     FROM Entrada_Produto"
    lSQL = lSQL & "    WHERE Empresa = " & g_empresa
    lSQL = lSQL & "      AND [Codigo do Produto] in " & lCodigosProdutos
    lSQL = lSQL & "      AND [Data da Entrada] = " & preparaData(lData)
    lSQL = lSQL & "      AND [Tipo da Entrada] <> " & 3 'Inventario
    lSQL = lSQL & " ORDER BY [Data da Entrada], [Numero do Documento], CFOP"
End Sub
Private Sub ImpDados()
    LoopTabelaGrupo
    If lPagina > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Movimentação do Estoque|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub LoopTabelaGrupo()
    'loop tabela Grupo
    lSQL = ""
    lSQL = lSQL & "   SELECT Codigo, Nome"
    lSQL = lSQL & "     FROM Grupo"
    If cbo_grupo.ItemData(cbo_grupo.ListIndex) <> 0 Then
        lSQL = lSQL & "    WHERE Codigo = " & Val(cbo_grupo.ItemData(cbo_grupo.ListIndex))
    End If
    lSQL = lSQL & " ORDER BY Nome, Codigo"
    'Abre RecordSet
    Set rsGrupo = New adodb.Recordset
    Set rsGrupo = Conectar.RsConexao(lSQL)
    
    If rsGrupo.RecordCount > 0 Then
        rsGrupo.MoveFirst
        Do Until rsGrupo.EOF
            Call LoopTabelaProduto(rsGrupo("Codigo").Value)
            rsGrupo.MoveNext
        Loop
    End If
    rsGrupo.Close
    Set rsGrupo = Nothing
End Sub
Private Sub LoopTabelaProduto(ByVal pGrupo As Integer)
    'loop tabela produto
    lSQL = ""
    lSQL = lSQL & "   SELECT Codigo, Nome"
    lSQL = lSQL & "     FROM Produto"
    lSQL = lSQL & "    WHERE [Codigo do Grupo] = " & pGrupo
    lSQL = lSQL & "      AND Inativo = " & preparaBooleano(False)
    'Início Select Tabela Entrada e Saida
    lSQL = lSQL & "      AND Codigo IN"
    lSQL = lSQL & " "
    lSQL = lSQL & " ("
    lSQL = lSQL & "   SELECT [Codigo do Produto] As Codigo2"
    lSQL = lSQL & "     FROM Entrada_Produto"
    lSQL = lSQL & "    WHERE Empresa = " & g_empresa
    lSQL = lSQL & "      AND [Data da Entrada] >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "      AND [Data da Entrada] <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & " GROUP BY [Codigo do Produto]"
    lSQL = lSQL & " "
    lSQL = lSQL & "    UNION"
    lSQL = lSQL & " "
    lSQL = lSQL & "   SELECT [Codigo do Produto2] As Codigo2"
    lSQL = lSQL & "     FROM Movimento_Lubrificante"
    lSQL = lSQL & "    WHERE Empresa = " & g_empresa
    lSQL = lSQL & "      AND Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "      AND Data <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & " GROUP BY [Codigo do Produto2]"
    lSQL = lSQL & " )"
    'Fim Select Tabela Entrada e Saida
    lSQL = lSQL & " ORDER BY Nome, Codigo"
    'Abre RecordSet
    Set rsProduto = New adodb.Recordset
    Set rsProduto = Conectar.RsConexao(lSQL)
    
    If rsProduto.RecordCount > 0 Then
        rsProduto.MoveFirst
        Do Until rsProduto.EOF
            If cbo_produto.ItemData(cbo_produto.ListIndex) = 0 Or cbo_produto.ItemData(cbo_produto.ListIndex) = rsProduto("Codigo").Value Then
                If lPagina = 0 Then
                    ImpCab
                End If
                'Le tabela auxiliar
                If chkCupomFiscal.Value = 1 Then
                    lSubAnterior = 0
                    lSubAtual = 0
                    If Estoque2.LocalizarCodigo(g_empresa, CDate(msk_data_i.Text) - 1, rsProduto("Codigo").Value) Then
                        lSubAnterior = Estoque2.Quantidade
                        lSubAtual = Estoque2.Quantidade
                    End If
                End If
                lSubEntrada = 0
                lSubSaida = 0
                lImprimiuCab2 = False
                lData = CDate(msk_data_i.Text)
                lImprimiuProduto = False
                Do Until lData > CDate(msk_data_f.Text)
                    Call LoopMovimentoEntrada(rsProduto("Codigo").Value)
                    Call LoopMovimentoSaida(rsProduto("Codigo").Value)
                    lData = lData + 1
                Loop
                If lSubAnterior <> 0 Or lSubEntrada <> 0 Or lSubSaida <> 0 Or lSubAtual Then
                    Call ImpSubTotal
                End If
            End If
            rsProduto.MoveNext
        Loop
    End If
    rsProduto.Close
    Set rsProduto = Nothing
End Sub
Private Sub LoopMovimentoEntrada(ByVal pCodigo As Long)
    'loop tabela Entrada_Produto
    lSQL = ""
    lSQL = lSQL & "   SELECT [Data da Entrada], [Numero do Documento], CFOP, Quantidade"
    lSQL = lSQL & "     FROM Entrada_Produto"
    lSQL = lSQL & "    WHERE Empresa = " & g_empresa
    lSQL = lSQL & "      AND [Codigo do Produto] = " & pCodigo
    lSQL = lSQL & "      AND [Data da Entrada] = " & preparaData(lData)
    lSQL = lSQL & "      AND [Tipo da Entrada] <> " & 3 'Inventario
    lSQL = lSQL & " ORDER BY [Data da Entrada], [Numero do Documento], CFOP"
    'Abre RecordSet
    Set rsEntradaProduto = New adodb.Recordset
    Set rsEntradaProduto = Conectar.RsConexao(lSQL)
    
    If rsEntradaProduto.RecordCount > 0 Then
        rsEntradaProduto.MoveFirst
        Do Until rsEntradaProduto.EOF
            If Not lImprimiuCab2 Then
                Call ImpCab2
            End If
            If lImprimiuProduto = False Then
                lImprimiuProduto = True
                Call ImpDet(rsProduto("Codigo").Value, rsProduto("Nome").Value, lSubAnterior)
            End If
            Cfop.Operacao = "+"
            If Cfop.LocalizarCodigo(rsEntradaProduto("CFOP").Value) = False Then
                MsgBox "CFOP inexistente!" + vbCrLf + "Cfop: " + rsEntradaProduto("CFOP").Value, vbCritical, "Erro de Integridade"
            End If
            Call ImpDet2(rsEntradaProduto("Data da Entrada").Value, rsEntradaProduto("Numero do Documento").Value, "ENTRADA P/COMPRA", rsEntradaProduto("Quantidade").Value, 0, Cfop.Operacao)
            rsEntradaProduto.MoveNext
        Loop
    End If
    rsEntradaProduto.Close
    Set rsEntradaProduto = Nothing
End Sub
Private Sub LoopRsMovimentoEntrada(ByVal pCodigo As Long)
    Dim xStrCondicao As String
    Dim rsEntradaProdutoFiltrada As New adodb.Recordset
    'loop RecordSet Entrada_Produto
    
    If rsEntradaProduto.RecordCount > 0 Then
        rsEntradaProduto.MoveFirst
        xStrCondicao = "[Codigo do Produto] = " & pCodigo
        xStrCondicao = xStrCondicao & " AND [Data da Entrada] = " & preparaData(lData)
        Set rsEntradaProdutoFiltrada = rsEntradaProduto.Clone()
        rsEntradaProdutoFiltrada.Filter = xStrCondicao
        
        
        If Not rsEntradaProdutoFiltrada.EOF Then
            Do Until rsEntradaProdutoFiltrada.EOF
                If Not lImprimiuCab2 Then
                    Call ImpCab2
                End If
                If lImprimiuProduto = False Then
                    lImprimiuProduto = True
                    Call ImpDet(rsProduto("Codigo").Value, rsProduto("Nome").Value, lSubAnterior)
                End If
                Cfop.Operacao = "+"
                If Cfop.LocalizarCodigo(rsEntradaProdutoFiltrada("CFOP").Value) = False Then
                    MsgBox "CFOP inexistente!" + vbCrLf + "Cfop: " + rsEntradaProdutoFiltrada("CFOP").Value, vbCritical, "Erro de Integridade"
                End If
                Call ImpDet2(rsEntradaProdutoFiltrada("Data da Entrada").Value, rsEntradaProdutoFiltrada("Numero do Documento").Value, "ENTRADA P/COMPRA", rsEntradaProdutoFiltrada("Quantidade").Value, 0, Cfop.Operacao)
                
                
                rsEntradaProdutoFiltrada.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub LoopRsMovimentoSaida(ByVal pCodigo As Long)
    Dim xStrCondicao As String
    Dim rsMovimentoSaidaFiltrada As New adodb.Recordset
    
    'loop RecordSet Movimento Saida
    
    If rsMovimentoLubrificante.RecordCount > 0 Then
        rsMovimentoLubrificante.MoveFirst
        xStrCondicao = "[Codigo do Produto2] = " & pCodigo
        xStrCondicao = xStrCondicao & " AND Data = " & preparaData(lData)
        Set rsMovimentoSaidaFiltrada = rsMovimentoLubrificante.Clone()
        rsMovimentoSaidaFiltrada.Filter = xStrCondicao
        
        
        If Not rsMovimentoSaidaFiltrada.EOF Then
            Do Until rsMovimentoSaidaFiltrada.EOF
                If Not lImprimiuCab2 Then
                    Call ImpCab2
                End If
                If lImprimiuProduto = False Then
                    lImprimiuProduto = True
                    Call ImpDet(rsProduto("Codigo").Value, rsProduto("Nome").Value, lSubAnterior)
                End If
                Call ImpDet2(rsMovimentoSaidaFiltrada("Data").Value, "Per." & rsMovimentoSaidaFiltrada("Periodo").Value, "SAIDA POR VENDAS", rsMovimentoSaidaFiltrada("Quantidade").Value, rsMovimentoSaidaFiltrada("Codigo do Funcionario").Value, "-")
                rsMovimentoSaidaFiltrada.MoveNext
            Loop
        End If
    End If
End Sub
Private Sub LoopMovimentoSaida(ByVal pCodigo As Long)
    'loop tabela Movimento_Nota_Fiscal_Saida
    lSQL = ""
    If chkCupomFiscal.Value = 0 Then
        lSQL = lSQL & "   SELECT Data, Periodo, Quantidade, [Codigo do Funcionario]"
        lSQL = lSQL & "     FROM Movimento_Lubrificante"
        lSQL = lSQL & "    WHERE Empresa = " & g_empresa
        lSQL = lSQL & "      AND [Codigo do Produto2] = " & pCodigo
        lSQL = lSQL & "      AND Data = " & preparaData(lData)
        lSQL = lSQL & " ORDER BY Data, Periodo"
    Else
        lSQL = lSQL & "   SELECT Data, Periodo, Quantidade, Operador AS [Codigo do Funcionario]"
        lSQL = lSQL & "     FROM Movimento_Cupom_Fiscal"
        lSQL = lSQL & "    WHERE Empresa = " & g_empresa
        lSQL = lSQL & "      AND [Codigo do Produto] = " & pCodigo
        lSQL = lSQL & "      AND Data = " & preparaData(lData)
        lSQL = lSQL & "      AND [Item Cancelado] = " & preparaBooleano(False)
        lSQL = lSQL & "      AND [Cupom Cancelado] = " & preparaBooleano(False)
        lSQL = lSQL & " ORDER BY Data, Periodo"
    End If
    'Abre RecordSet
    Set rsMovimentoLubrificante = New adodb.Recordset
    Set rsMovimentoLubrificante = Conectar.RsConexao(lSQL)
    
    If rsMovimentoLubrificante.RecordCount > 0 Then
        rsMovimentoLubrificante.MoveFirst
        Do Until rsMovimentoLubrificante.EOF
            If Not lImprimiuCab2 Then
                Call ImpCab2
            End If
            If lImprimiuProduto = False Then
                lImprimiuProduto = True
                Call ImpDet(rsProduto("Codigo").Value, rsProduto("Nome").Value, lSubAnterior)
            End If
            Call ImpDet2(rsMovimentoLubrificante("Data").Value, "Per." & rsMovimentoLubrificante("Periodo").Value, "SAIDA POR VENDAS", rsMovimentoLubrificante("Quantidade").Value, rsMovimentoLubrificante("Codigo do Funcionario").Value, "-")
            rsMovimentoLubrificante.MoveNext
        Loop
    End If
    rsMovimentoLubrificante.Close
    Set rsMovimentoLubrificante = Nothing
End Sub
Private Sub ImpDet(ByVal pCodigo As Long, ByVal pNome As String, ByVal pQuantidade As Currency)
    Dim xLinha As String
    Dim i As Integer
    
    If lLinha >= 58 Then
        lGrupo = 0
        xLinha = "+----------+------+-----------------+------------+------------+------------+---+"
        Mid(xLinha, 26, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    'If pCodigo_grupo <> lGrupo Then
    '    Call ImpSubTotal(pCodigo_grupo, pNome_grupo)
    'End If
              '         1         2         3         4         5         6         7         8
              '12345678901234567890123456789012345678901234567890123456789012345678901234567890
    If lLinha = 0 Then
        xLinha = "+------------------------------------------------------------------------------+"
    Else
        xLinha = "+----------+------+-----------------+------------+------------+------------+---+"
    End If
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = True"
    xLinha = "| PRODUTO.:                                                                    |"
    i = Len(Format(pCodigo, "#,000"))
    Mid(xLinha, 13 + 5 - i, i) = Format(pCodigo, "#,000")
    Mid(xLinha, 19, 40) = pNome
    i = Len(Format(pQuantidade, "########0.00"))
    Mid(xLinha, vbInformation + 12 - i, i) = Format(pQuantidade, "########0.00")
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    lLinha = lLinha + 2
End Sub
Private Sub ImpDet2(ByVal pData As Date, ByVal pNumeroNota As String, ByVal pDescricao As String, ByVal pQuantidade As Currency, ByVal pCodigoVendedor As Integer, ByVal pOperacao As String)
    Dim xLinha As String
    Dim i As Integer

    If lLinha >= 58 Then
        lGrupo = 0
        xLinha = "+----------+------+-----------------+------------+------------+------------+---+"
        Mid(xLinha, 5, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        Call ImpCab
    End If
    If lLinha = 0 Then
        Call ImpDet(rsProduto("Codigo").Value, rsProduto("Nome").Value, lSubAtual)
        Call ImpCab2
    End If
              '         1         2         3         4         5         6         7         8
              '12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "|          |      |                 |            |            |            |   |"
    Mid(xLinha, 2, 10) = Format(pData, "dd/mm/yyyy")
    i = Len(Format(pNumeroNota, "#####0"))
    Mid(xLinha, 13 + 6 - i, i) = Format(pNumeroNota, "#####0")
    Mid(xLinha, 20, 17) = pDescricao
    If pOperacao = "+" Then
        i = Len(Format(pQuantidade, "########0.00"))
        Mid(xLinha, 38 + 12 - i, i) = Format(pQuantidade, "########0.00")
        lSubEntrada = lSubEntrada + pQuantidade
        lSubAtual = lSubAtual + pQuantidade
    ElseIf pOperacao = "-" Then
        i = Len(Format(pQuantidade, "########0.00"))
        Mid(xLinha, 51 + 12 - i, i) = Format(pQuantidade, "########0.00")
        lSubSaida = lSubSaida + pQuantidade
        lSubAtual = lSubAtual - pQuantidade
    End If
    i = Len(Format(lSubAtual, "########0.00"))
    Mid(xLinha, vbInformation + 12 - i, i) = Format(lSubAtual, "########0.00")
    If pCodigoVendedor > 0 Then
        i = Len(Format(pCodigoVendedor, "##0"))
        Mid(xLinha, 77 + 3 - i, i) = Format(pCodigoVendedor, "##0")
    End If
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpSubTotal()
    Dim xLinha As String
    Dim i As Integer
    If lLinha >= 58 Then
        xLinha = "+----------+------+-----------------+------------+------------+------------+---+"
        Mid(xLinha, 5, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    xLinha = "+----------+------+-----------------+------------+------------+------------+---+"
    BioImprime "@Printer.Print " & xLinha
              '         1         2         3         4         5         6         7         8
              '12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "| TOTAL DO PRODUTO|                 |            |            |            |   |"
    i = Len(Format(lSubAnterior, "########0.00"))
    Mid(xLinha, 20 + 12 - i, i) = Format(lSubAnterior, "########0.00")
    If lSubEntrada > 0 Then
        i = Len(Format(lSubEntrada, "########0.00"))
        Mid(xLinha, 38 + 12 - i, i) = Format(lSubEntrada, "########0.00")
    End If
    If lSubSaida > 0 Then
        i = Len(Format(lSubSaida, "########0.00"))
        Mid(xLinha, 51 + 12 - i, i) = Format(lSubSaida, "########0.00")
    End If
    i = Len(Format(lSubAtual, "########0.00"))
    Mid(xLinha, vbInformation + 12 - i, i) = Format(lSubAtual, "########0.00")
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 2
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    xLinha = "+----------+------+-----------------+------------+------------+------------+---+"
    Mid(xLinha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpCab()
    Dim xLinha As String
    Dim i As Integer
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
    BioImprime "@@Printer.Print " & Chr(34) & " " & Chr(34)
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.CurrentY = 0"
              '         1         2         3         4         5         6         7         8
              '12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "+------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = True"
    xLinha = "|                                                                  Página: ___ |"
    Mid(xLinha, 3, 40) = g_nome_empresa
    Mid(xLinha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    xLinha = "| MOVIMENTACAO DE ESTOQUE                                   CIDADE, __/__/____ |"
    i = Len(Trim(g_cidade_empresa))
    Mid(xLinha, 37 + 30 - i, i) = Trim(g_cidade_empresa)
    Mid(xLinha, 69, 10) = msk_data
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| REFERENTE A.......: __/__/____ A __/__/____                                  |"
    Mid(xLinha, 23, 10) = msk_data_i
    Mid(xLinha, 36, 10) = msk_data_f
    BioImprime "@Printer.Print " & xLinha
End Sub
Private Sub ImpCab2()
    Dim xLinha As String
    If lLinha >= 58 Then
        lGrupo = 0
        xLinha = "+----------+------+-----------------+------------+------------+------------+---+"
        Mid(xLinha, 5, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
'    If lLinha = 0 Then
'        Call ImpDet(rsProduto("Codigo").Value, rsProduto("Nome").Value, lSubAtual)
'    End If
    xLinha = "+----------+------+-----------------+------------+------------+------------+---+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|   DATA   |N.NOTA|DESC. MOVIMENTO  |  ENTRADAS  |   SAIDAS   |   SALDOS   |VEN|"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+----------+------+-----------------+------------+------------+------------+---+"
    'BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 3
    lImprimiuCab2 = True
End Sub
Private Sub cbo_grupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_produto.SetFocus
    End If
End Sub
Private Sub cbo_produto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cbo_grupo.SetFocus
    Else
        msk_data = RetiraGString(1)
        msk_data_i.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_data_f_Click()
    g_string = msk_data_f
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
    Else
        msk_data_f = RetiraGString(1)
    End If
    g_string = " "
    cbo_grupo.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cbo_grupo.SetFocus
    Else
        msk_data_i = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = " "
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
    ElseIf cbo_grupo.ListIndex = -1 Then
        MsgBox "Escolha o grupo.", vbInformation, "Atenção!"
        cbo_grupo.SetFocus
    ElseIf cbo_produto.ListIndex = -1 Then
        MsgBox "Escolha o produto.", vbInformation, "Atenção!"
        cbo_produto.SetFocus
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
        msk_data_i.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def, "dd/mm/yyyy")
        cbo_grupo.ListIndex = 0
        cbo_produto.ListIndex = 0
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
    PreencheCboGrupo
    PreencheCboProduto
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
        cbo_produto.SetFocus
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
