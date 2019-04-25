VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form EmissaoCalculoPrecoMedio 
   Caption         =   "Emissão de Cálculo de Preço Médio"
   ClientHeight    =   2220
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "EmissaoCalculoPrecoMedio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "EmissaoCalculoPrecoMedio.frx":030A
   ScaleHeight     =   2220
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "EmissaoCalculoPrecoMedio.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Visualiza a medida de tanque."
      Top             =   1260
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "EmissaoCalculoPrecoMedio.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprime a medida de tanque."
      Top             =   1260
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "EmissaoCalculoPrecoMedio.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1260
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CheckBox chkVerificaPreco 
         Caption         =   "&Diverg&encia de Preços"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   660
         Width           =   2235
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "EmissaoCalculoPrecoMedio.frx":4706
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
      Begin VB.Label Label1 
         Caption         =   "Verifica"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
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
      Top             =   1500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "EmissaoCalculoPrecoMedio"
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
Dim lNomeEmpresa As String
Dim lSQL As String
Dim lSubTotalAnterior As Currency
Dim lSubTotalAtual As Currency
Dim lSubTotalEntrada As Currency
Dim lSubTotalVenda As Currency
Dim lSubTotalPerdaSobra As Currency
Dim lTotalAtual As Currency
Dim lQtdEmpresa As Integer

Private Bomba As New CadastroDLL.cBomba
Private EntradaCombustivel As New CadastroDLL.cEntradaCombustivel
Private Estoque As New CadastroDLL.cEstoque
Private MovimentoBomba As New CadastroDLL.cMovimentoBomba
Private MovimentoCupomFiscal As New CadastroDLL.cMovimentoCupomFiscal
Private Produto As New CadastroDLL.cProduto

Private rsEmpresaCombustivel As New adodb.Recordset
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Bomba = Nothing
    Set EntradaCombustivel = Nothing
    Set Estoque = Nothing
    Set MovimentoBomba = Nothing
    Set MovimentoCupomFiscal = Nothing
    Set Produto = Nothing
    Set rsEmpresaCombustivel = Nothing
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lSubTotalAnterior = 0
    lSubTotalAtual = 0
    lSubTotalEntrada = 0
    lSubTotalVenda = 0
    lSubTotalPerdaSobra = 0
    lTotalAtual = 0
    lQtdEmpresa = 0
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Empresas.Nome AS NomeEmpresa, Empresas.Codigo AS CodigoEmpresa, Combustivel.Nome, Combustivel.Codigo AS CodigoCombustivel, Combustivel.[Preco Medio]"
    lSQL = lSQL & "  FROM Empresas, Combustivel"
    lSQL = lSQL & " WHERE Combustivel.Empresa = Empresas.Codigo"
    lSQL = lSQL & " ORDER BY Empresas.Nome, Combustivel.Nome"
    
    'Abre RecordSet
    Set rsEmpresaCombustivel = New adodb.Recordset
    Set rsEmpresaCombustivel = Conectar.RsConexao(lSQL)
    
    
    'Verifica dados
    If rsEmpresaCombustivel.RecordCount > 0 Then
        ImpDados
    End If
    If rsEmpresaCombustivel.State = 1 Then
        rsEmpresaCombustivel.Close
    End If
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    Dim xLinha As String
    Dim xNomeArquivo As String
    Dim retval As Long
    
    'loop dos Dados
    Do Until rsEmpresaCombustivel.EOF
        If lPagina = 0 Then
            ImpCab
            lNomeEmpresa = rsEmpresaCombustivel("NomeEmpresa").Value
            ImpCabEmpresa (False)
            ImpCabEmpresa2
        End If
        If rsEmpresaCombustivel("NomeEmpresa").Value <> lNomeEmpresa Then
            ImpSubTotal
            lNomeEmpresa = rsEmpresaCombustivel("NomeEmpresa").Value
            ImpCabEmpresa (False)
            ImpCabEmpresa2
        End If
        If lLinha >= 75 Then
            xLinha = "+--------------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+"
            Mid(xLinha, 5, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & xLinha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        ImpDet
        rsEmpresaCombustivel.MoveNext
    Loop
    If lPagina > 0 Then
        ImpSubTotal
        ImpTotal
    End If
    
    lNomeEmpresa = ""
    'loop dos Dados II (para verificar erro nos preços de venda)
    If chkVerificaPreco.Value = 1 Then
        BioImprime "@@Printer.NewPage"
        rsEmpresaCombustivel.MoveFirst
        Do Until rsEmpresaCombustivel.EOF
            If lLinha >= 75 Then
                xLinha = "+--------------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+"
                Mid(xLinha, 5, 22) = " Cerrado Informática. "
                BioImprime "@Printer.Print " & xLinha
                BioImprime "@@Printer.NewPage"
                ImpCab
            End If
            ImpDetDiferencaPreco
            rsEmpresaCombustivel.MoveNext
        Loop
    End If
    If lPagina > 0 Then
        If lNomeEmpresa <> "" Then
            ImpSubTotal
            ImpTotal
        End If
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Medição de Combustíveis|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub ImpCabEmpresa(ByVal pDiferencaPreco As Boolean)
    Dim xLinha As String
    Dim i As Integer
    '                  1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    xLinha = "| DADOS ABAIXO SE REFERE A EMPRESA.:                                                                                                    |"
    Mid(xLinha, 38, 40) = rsEmpresaCombustivel("NomeEmpresa").Value
    If pDiferencaPreco Then
        Mid(xLinha, 118, 18) = "preco maior do dia"
    End If
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpCabEmpresa2()
    '                  1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    BioImprime "@Printer.Print " & "+--------------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+"
    BioImprime "@Printer.Print " & "|COMBUSTIVEL   | DATA ULT.|  NUMERO  | QTD.  EM |   TOTAL  |   CUSTO  | PCO.VENDA|   LUCRO  |   LUCRO  | PCO.VENDA|   LUCRO  | PCO.VENDA|"
    BioImprime "@Printer.Print " & "|              |  ENTRADA |  DA NOTA |  LITROS  | DA  NOTA | UNITARIO | UNITARIO |   BRUTO  | DESEJADO | DESEJADO |   MEDIO  |   MEDIO  |"
    BioImprime "@Printer.Print " & "+--------------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+"
    lLinha = lLinha + 4
End Sub
Private Sub ImpDet()
    Dim xLinha As String
    Dim i As Integer
    Dim xValorProduto As Currency
    Dim xLucroDesejado As Currency
    Dim xLucroBruto As Currency
    Dim xPrecoVendaDesejado As Currency
    Dim xLucroMedio As Currency
    Dim xPrecoMedio As Currency
    '                  1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
             '+---------------------------------------------------------------------------------------------------------------------------------------+
             '| DADOS ABAIXO SE REFERE A EMPRESA.:                                                                                                    |
             '+--------------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+
             '|COMBUSTIVEL   | DATA ULT.|  NUMERO  | QTD.  EM |   TOTAL  |   CUSTO  | PCO.VENDA|   LUCRO  |   LUCRO  | PCO.VENDA|   LUCRO  | PCO.VENDA|
             '|              |  ENTRADA |  DA NOTA |  LITROS  | DA  NOTA | UNITARIO | UNITARIO |   BRUTO  | DESEJADO | DESEJADO |   MEDIO  |   MEDIO  |
             '+--------------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+
             '|12345678901234|12/12/2009|1234567890|1234567890| 234567890| 234567890| 234567890| 234567890| 234567890| 234567890| 234567890| 234567890|
             '+--------------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+
    xValorProduto = 0
    xPrecoVendaDesejado = 0
    xLucroDesejado = 0
    xLucroMedio = 0
    xPrecoMedio = 0
    If rsEmpresaCombustivel("CodigoCombustivel").Value = "A " Then
        xLucroDesejado = 0.2
    ElseIf rsEmpresaCombustivel("CodigoCombustivel").Value = "D " Then
        xLucroDesejado = 0.05
    ElseIf rsEmpresaCombustivel("CodigoCombustivel").Value = "G " Then
        xLucroDesejado = 0.2
    End If
    If Bomba.LocalizarTipoCombustivel(rsEmpresaCombustivel("CodigoEmpresa").Value, rsEmpresaCombustivel("CodigoCombustivel").Value) Then
        If Estoque.LocalizarCodigo(rsEmpresaCombustivel("CodigoEmpresa").Value, Bomba.CodigoProduto) Then
            xValorProduto = Estoque.PrecoVenda
        End If
    End If
    xLinha = "|              |          |          |          |          |          |          |          |          |          |          |          |"
    Mid(xLinha, 2, 14) = rsEmpresaCombustivel("Nome").Value
    If EntradaCombustivel.LocalizarUltimoCombustivel(rsEmpresaCombustivel("CodigoEmpresa").Value, CDate(msk_data.Text), rsEmpresaCombustivel("CodigoCombustivel").Value) Then
        Mid(xLinha, 17, 10) = Format(EntradaCombustivel.Data, "dd/MM/yyyy")
        Mid(xLinha, 28, 10) = EntradaCombustivel.NumeroNota
        i = Len(Format(EntradaCombustivel.Quantidade, "##,###,##0"))
        Mid(xLinha, 39 + 10 - i, i) = Format(EntradaCombustivel.Quantidade, "##,###,##0")
        i = Len(Format(EntradaCombustivel.ValorEntrada, "###,##0.00"))
        Mid(xLinha, 50 + 10 - i, i) = Format(EntradaCombustivel.ValorEntrada, "###,##0.00")
        i = Len(Format(EntradaCombustivel.ValorLitro, "#,##0.0000"))
        Mid(xLinha, 61 + 10 - i, i) = Format(EntradaCombustivel.ValorLitro, "#,##0.0000")
        i = Len(Format(xValorProduto, "#,##0.0000"))
        Mid(xLinha, 72 + 10 - i, i) = Format(xValorProduto, "#,##0.0000")
        xLucroBruto = xValorProduto - EntradaCombustivel.ValorLitro
        i = Len(Format(xLucroBruto, "#,##0.0000"))
        Mid(xLinha, 83 + 10 - i, i) = Format(xLucroBruto, "#,##0.0000")
        xPrecoVendaDesejado = EntradaCombustivel.ValorLitro + xLucroDesejado
        If xLucroBruto > xLucroDesejado Then
            xLucroMedio = xLucroDesejado
        Else
            xLucroMedio = xLucroBruto / 2
        End If
        xPrecoMedio = EntradaCombustivel.ValorLitro + xLucroMedio
    End If
    i = Len(Format(xLucroDesejado, "###,##0.00"))
    Mid(xLinha, 94 + 10 - i, i) = Format(xLucroDesejado, "###,##0.00")
    i = Len(Format(xPrecoVendaDesejado, "#,##0.0000"))
    Mid(xLinha, 105 + 10 - i, i) = Format(xPrecoVendaDesejado, "#,##0.0000")
    
    i = Len(Format(xLucroMedio, "#,##0.0000"))
    Mid(xLinha, 116 + 10 - i, i) = Format(xLucroMedio, "#,##0.0000")
    
    i = Len(Format(xPrecoMedio, "#,##0.0000"))
    Mid(xLinha, 127 + 10 - i, i) = Format(xPrecoMedio, "#,##0.0000")

    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpDetDiferencaPreco()
    Dim xLinha As String
    Dim i As Integer
    Dim xPrecoBomba As Currency
    Dim xPrecoProduto As Currency
    Dim xPrecoMovBomba As Currency
    Dim xPrecoCupom As Currency
    Dim xDataMovBomba As String
    Dim xDataCupom As String
    '                  1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
             '+---------------------------------------------------------------------------------------------------------------------------------------+
             '| DADOS ABAIXO SE REFERE A EMPRESA.:                                                                                                    |
             '+--------------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+
             '|COMBUSTIVEL   | DATA ULT.|  NUMERO  | QTD.  EM |   TOTAL  |   CUSTO  | PCO.VENDA|   LUCRO  |   LUCRO  | PCO.VENDA|   LUCRO  | PCO.VENDA|
             '|              |  ENTRADA |  DA NOTA |  LITROS  | DA  NOTA | UNITARIO | UNITARIO |   BRUTO  | DESEJADO | DESEJADO |   MEDIO  |   MEDIO  |
             '+--------------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+
             '|              |          |     PRECOS DE VENDA | MOV.BOMBA|          | CAD.BOMBA|          |CAD.PRODUT|          |ECF: 99/99|          |
             '+--------------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+
    
    xPrecoBomba = 0
    xPrecoProduto = 0
    xPrecoMovBomba = 0
    xPrecoCupom = 0
    xDataMovBomba = ""
    xDataCupom = ""
    If Bomba.LocalizarTipoCombustivel(rsEmpresaCombustivel("CodigoEmpresa").Value, rsEmpresaCombustivel("CodigoCombustivel").Value) Then
        xPrecoBomba = Bomba.PrecoVenda
        If Estoque.LocalizarCodigo(rsEmpresaCombustivel("CodigoEmpresa").Value, Bomba.CodigoProduto) Then
            xPrecoProduto = Estoque.PrecoVenda
        End If
        If MovimentoCupomFiscal.LocalizarProdutoAnteriorData(rsEmpresaCombustivel("CodigoEmpresa").Value, CDate(msk_data.Text), Bomba.CodigoProduto, False) Then
            xPrecoCupom = MovimentoCupomFiscal.ValorUnitario
            xDataCupom = Format(MovimentoCupomFiscal.Data, "dd/MM")
        End If
    End If
    If MovimentoBomba.LocalizarCombAnteriorData(rsEmpresaCombustivel("CodigoEmpresa").Value, CDate(msk_data.Text), rsEmpresaCombustivel("CodigoCombustivel").Value) Then
        xPrecoMovBomba = MovimentoBomba.PrecoVenda
        xDataMovBomba = Format(MovimentoBomba.Data, "dd/MM/yyyy")
    End If
    

    If xPrecoBomba <> xPrecoProduto Or xPrecoBomba <> xPrecoMovBomba Or xPrecoBomba <> xPrecoCupom Then
        If lNomeEmpresa = "" Then
            ImpCab
        End If
        If rsEmpresaCombustivel("NomeEmpresa").Value <> lNomeEmpresa Then
            If lNomeEmpresa <> "" Then
                BioImprime "@Printer.Print " & "+--------------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+"
                lLinha = lLinha + 1
            End If
            lNomeEmpresa = rsEmpresaCombustivel("NomeEmpresa").Value
            ImpCabEmpresa (True)
        End If
        If lLinha >= 75 Then
            xLinha = "+--------------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+"
            Mid(xLinha, 5, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & xLinha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        xLinha = "|              |          |     PRECOS DE VENDA | MOV.BOMBA|          | CAD.BOMBA|          |CAD.PRODUT|          |ECF: 99/99|          |"
        Mid(xLinha, 2, 14) = rsEmpresaCombustivel("Nome").Value
        Mid(xLinha, 17, 10) = xDataMovBomba
        i = Len(Format(xPrecoMovBomba, "#,##0.0000"))
        Mid(xLinha, 61 + 10 - i, i) = Format(xPrecoMovBomba, "#,##0.0000")
        If xPrecoMovBomba <> xPrecoBomba Then
            Mid(xLinha, 61, 2) = "**"
        End If
        i = Len(Format(xPrecoBomba, "#,##0.0000"))
        Mid(xLinha, 83 + 10 - i, i) = Format(xPrecoBomba, "#,##0.0000")
        i = Len(Format(xPrecoProduto, "#,##0.0000"))
        Mid(xLinha, 105 + 10 - i, i) = Format(xPrecoProduto, "#,##0.0000")
        If xPrecoProduto <> xPrecoBomba Then
            Mid(xLinha, 105, 2) = "**"
        End If
        Mid(xLinha, 121, 5) = xDataCupom
        i = Len(Format(xPrecoCupom, "#,##0.0000"))
        Mid(xLinha, 127 + 10 - i, i) = Format(xPrecoCupom, "#,##0.0000")
        If xPrecoCupom <> xPrecoBomba Then
            Mid(xLinha, 127, 2) = "**"
        End If
        
        
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
    End If
End Sub
Private Sub ImpSubTotal()
    Dim xLinha As String
        
    lQtdEmpresa = lQtdEmpresa + 1
    BioImprime "@Printer.Print " & "+--------------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+"
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    
    xLinha = "+--------------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+----------+"
    Mid(xLinha, 5, 22) = " Cerrado Informática. "
    Mid(xLinha, 55, 2) = Format(lQtdEmpresa, "00")
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
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
    xLinha = "| RELAÇÃO DOS CÁLCULOS DE PREÇO MÉDIO                       CIDADE, __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+---------------------------------------------------------------------------------------------------------------------------------------+"
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data.Text = RetiraGString(1)
    cmd_visualizar.SetFocus
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
        cmd_visualizar.SetFocus
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
    If g_nome_usuario = "L.M.C." Then
        Me.Caption = Me.Caption & " - LMC"
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel_LMC"
        MovimentoBomba.NomeTabela = "Movimento_Bomba_LMC"
    Else
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel"
        MovimentoBomba.NomeTabela = "Movimento_Bomba"
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_GotFocus()
    msk_data.SelStart = 0
    msk_data.SelLength = 2
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
