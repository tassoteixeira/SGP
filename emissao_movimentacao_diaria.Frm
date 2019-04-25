VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_movimentacao_diaria 
   Caption         =   "Emissão da Movimentação Diária"
   ClientHeight    =   2070
   ClientLeft      =   1965
   ClientTop       =   3705
   ClientWidth     =   5325
   Icon            =   "emissao_movimentacao_diaria.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "emissao_movimentacao_diaria.frx":030A
   ScaleHeight     =   2070
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   3840
      Picture         =   "emissao_movimentacao_diaria.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Visualiza a medida de tanque."
      Top             =   1080
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   6135
      Left            =   180
      ScaleHeight     =   6075
      ScaleWidth      =   6615
      TabIndex        =   6
      Top             =   2100
      Visible         =   0   'False
      Width           =   6675
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   960
      Picture         =   "emissao_movimentacao_diaria.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprime a Movimentação Diária"
      Top             =   1080
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2220
      Picture         =   "emissao_movimentacao_diaria.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1080
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "emissao_movimentacao_diaria.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   300
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_movimentacao_diaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Início de variáveis padrão para relatório
Dim lColuna(0 To 11) As Currency
Dim lLinha(0 To 17) As Currency
Dim lColunaI As Currency
Dim lLinhaI As Currency
Dim lLocal As Integer

Dim lEncerranteI(0 To 6) As Currency
Dim lEncerranteF(0 To 6) As Currency
Dim lQuantidade(0 To 6) As Currency
Dim lPrecoCusto(0 To 6) As Currency
Dim lPrecoVenda(0 To 6) As Currency
Dim lValorVendaDia(0 To 6) As Currency
Dim lValorVendaAcumulado(0 To 6) As Currency
Dim lPrevisaoVendaQtd(0 To 6) As Currency
Dim lPrevisaoVendaValor(0 To 6) As Currency
Dim lReceitaBrutaEstimada(0 To 6) As Currency
Dim lPercentual As Currency
Dim lMedidaA As Currency
Dim lMedidaG As Currency
Dim lQtdDiasMes As Integer

Dim lR As Integer
Dim lG As Integer
Dim lB As Integer
Dim lSQl As String
Private rsTabela As New adodb.Recordset
Private Sub ZeraVariaveis()
    lLocal = 1
    lR = 0
    lG = 0
    lB = 0
End Sub
Private Function cp(ByVal xCentimetros As Currency) As Currency
    cp = xCentimetros / 2.56
End Function
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
End Sub
Private Sub ImpDados()
    ZeraVariaveis
    'seleciona medidas para centímetros
    Printer.ColorMode = 2
    Printer.ScaleMode = 7
    Printer.PrintQuality = vbPRPQMedium
    Printer.PaperSize = vbPRPSLegal
    Printer.Orientation = 2
    Printer.FontName = "Arial"
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    Printer.FontBold = True
    
    'teste para imprimir letra correta
    Printer.FontBold = False
    ImprimeTexto "  ", lColuna(0), lColuna(1), 0, lLocal
    Printer.FontBold = True
    Printer.ForeColor = RGB(lR, lG, lB)
    
    ImprimeGradeA
    ImprimeGradeB
    ImprimeGradeC
    ImprimeGradeD
    ImprimeGradeE
    
    Printer.EndDoc
    
End Sub
Private Sub ImprimeGradeA()
    Dim i As Integer
    Dim i2 As Integer
    For i = 0 To 6
        lEncerranteI(i) = 0
        lEncerranteF(i) = 0
        lQuantidade(i) = 0
        lPrecoCusto(i) = 0
    Next
    'rsTabela.Sort = "Data, Periodo, [Codigo da Bomba]"
    'rsTabela.Filter = "Data = " & Format(msk_data.Text, "mm/dd/yyyy")
    Do Until rsTabela.EOF
        i2 = rsTabela("Codigo da Bomba").Value
        If lEncerranteI(i2) = 0 Then
            lEncerranteI(i2) = rsTabela("Abertura").Value
        End If
        lEncerranteF(i2) = rsTabela("Encerrante").Value
        lPrecoCusto(i2) = rsTabela("Preco de Custo").Value
        lQuantidade(i2) = lQuantidade(i2) + rsTabela("Quantidade da Saida").Value
        rsTabela.MoveNext
    Loop
    
    
    'Le Medição de Combustível
    Set rsTabela = Nothing
    lSQl = ""
    lSQl = lSQl & "SELECT Tipo_Combustivel, Tanque_1"
    lSQl = lSQl & "  FROM Medicao_Combustivel"
    lSQl = lSQl & " WHERE Empresa = " & g_empresa
    lSQl = lSQl & "   AND Data = #" & Format(CDate(msk_data.Text) + 1, "mm/dd/yyyy") & "#"
    lSQl = lSQl & " ORDER BY Tipo_Combustivel"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQl)
    Do Until rsTabela.EOF
        If Trim(rsTabela("Tipo_Combustivel").Value) = "A" Then
            lMedidaA = rsTabela("Tanque_1").Value
        ElseIf Trim(rsTabela("Tipo_Combustivel").Value) = "G" Then
            lMedidaG = rsTabela("Tanque_1").Value
        End If
        rsTabela.MoveNext
    Loop
    
    
    'Coordenada inicial da linha da grade A (em centimetros)
    lLinhaI = 0.3
    'Coordenada inicial da coluna da grade A (em centimetros)
    lColunaI = 3.8
    
    'Coordenadas das colunas da grade A (em centimetros)
    lColuna(0) = lColunaI + 0
    lColuna(1) = lColunaI + 2.2
    lColuna(2) = lColunaI + 3.9
    lColuna(3) = lColunaI + 5.6
    lColuna(4) = lColunaI + 7.3
    lColuna(5) = lColunaI + 9.6
    lColuna(6) = lColunaI + 11.44
    lColuna(7) = lColunaI + 13.28
    lColuna(8) = lColunaI + 15.12
    lColuna(9) = lColunaI + 16.96
    lColuna(10) = lColunaI + 18.8
    lColuna(11) = lColunaI + 20.5
    
    'Coordenadas das Linhas da grade A (em centimetros)
    lLinha(0) = lLinhaI + 0
    lLinha(1) = lLinhaI + 1
    lLinha(2) = lLinhaI + 1.62
    lLinha(3) = lLinhaI + 1.93
    lLinha(4) = lLinhaI + 2.24
    lLinha(5) = lLinhaI + 2.55
    lLinha(6) = lLinhaI + 2.86
    lLinha(7) = lLinhaI + 3.17
    lLinha(8) = lLinhaI + 3.48
    
    
    
    'Imprime os dados do cabeçalho do relatório
    Printer.DrawWidth = 6
    Printer.Line (lColuna(0), lLinha(0))-(lColuna(11), lLinha(0)), RGB(lR, lG, lB)
    Printer.FontSize = 12
    Printer.FontBold = True
    ImprimeCentralizado "MOVIMENTAÇÃO DIÁRIA " & Trim(UCase(g_nome_empresa)) & " - " & UCase(Format(CDate(msk_data.Text), "mmmm")) & "/" & Format(CDate(msk_data.Text), "yyyy"), lColuna(0), lColuna(10), lLinha(0) + 0.3, lLocal
    ImprimeCentralizado Format(Format(CDate(msk_data.Text), "dd"), "00"), lColuna(10), lColuna(11), lLinha(0) + 0.3, lLocal
    Printer.FontBold = False
    Printer.Line (lColuna(0), lLinha(1))-(lColuna(11), lLinha(1)), RGB(lR, lG, lB)
    
    'Imprime os dados do cabeçalho da grade A
    Printer.FontSize = 6
    ImprimeCentralizado "Produto", lColuna(0), lColuna(1), lLinha(1) + 0.21, lLocal
    ImprimeCentralizado "Encerrante Dia", lColuna(1), lColuna(2), lLinha(1) + 0.1, lLocal
    ImprimeCentralizado "Anterior", lColuna(1), lColuna(2), lLinha(1) + 0.36, lLocal
    ImprimeCentralizado "Iniciante", lColuna(2), lColuna(3), lLinha(1) + 0.21, lLocal
    ImprimeCentralizado "Volume Saída", lColuna(3), lColuna(4), lLinha(1) + 0.21, lLocal
    ImprimeCentralizado "Venda Dia", lColuna(4), lColuna(5), lLinha(1) + 0.21, lLocal
    ImprimeCentralizado "Estoque", lColuna(5), lColuna(6), lLinha(1) + 0.1, lLocal
    ImprimeCentralizado "Medição", lColuna(5), lColuna(6), lLinha(1) + 0.36, lLocal
    ImprimeCentralizado "Diferenças de", lColuna(6), lColuna(7), lLinha(1) + 0.1, lLocal
    ImprimeCentralizado "Medição", lColuna(6), lColuna(7), lLinha(1) + 0.36, lLocal
    ImprimeCentralizado "Compra Dia", lColuna(7), lColuna(8), lLinha(1) + 0.21, lLocal
    ImprimeCentralizado "Preço Compra", lColuna(8), lColuna(9), lLinha(1) + 0.21, lLocal
    ImprimeCentralizado "Previsão", lColuna(9), lColuna(10), lLinha(1) + 0.1, lLocal
    ImprimeCentralizado "Compra", lColuna(9), lColuna(10), lLinha(1) + 0.36, lLocal
    ImprimeCentralizado "Retorno", lColuna(10), lColuna(11), lLinha(1) + 0.21, lLocal
    
    'Imprime linha horizontal da grade A
    Printer.Line (lColuna(0), lLinha(2))-(lColuna(11), lLinha(2)), RGB(lR, lG, lB)
    
    'Imprime 1ª linha de dados da grade A
    ImprimeTexto "Gasolina C", lColuna(0), lColuna(1), lLinha(2) + 0.07, lLocal
    ImprimeValor Format(lEncerranteI(1), "#,###,##0.0"), lColuna(1), lColuna(2), lLinha(2) + 0.07, lLocal
    ImprimeValor Format(lEncerranteF(1), "#,###,##0.0"), lColuna(2), lColuna(3), lLinha(2) + 0.07, lLocal
    ImprimeValor Format(lQuantidade(1), "#,###,##0.0"), lColuna(3), lColuna(4), lLinha(2) + 0.07, lLocal
    ImprimeValor "0,0", lColuna(10), lColuna(11), lLinha(2) + 0.07, lLocal
    
    'Imprime linha horizontal da grade A
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(3))-(lColuna(4), lLinha(3)), RGB(lR, lG, lB)
    Printer.Line (lColuna(10), lLinha(3))-(lColuna(11), lLinha(3)), RGB(lR, lG, lB)
    
    'Imprime 2ª linha de dados da grade A
    ImprimeTexto "Gasolina C", lColuna(0), lColuna(1), lLinha(3) + 0.07, lLocal
    ImprimeValor Format(lEncerranteI(2), "#,###,##0.0"), lColuna(1), lColuna(2), lLinha(3) + 0.07, lLocal
    ImprimeValor Format(lEncerranteF(2), "#,###,##0.0"), lColuna(2), lColuna(3), lLinha(3) + 0.07, lLocal
    ImprimeValor Format(lQuantidade(2), "#,###,##0.0"), lColuna(3), lColuna(4), lLinha(3) + 0.07, lLocal
    ImprimeValor Format(lQuantidade(1) + lQuantidade(2), "#,###,##0.0"), lColuna(4), lColuna(5), lLinha(2) + 0.37, lLocal
    ImprimeTexto "R$", lColuna(8), lColuna(9), lLinha(2) + 0.37, lLocal
    ImprimeValor Format(lPrecoCusto(2), "###,##0.0000"), lColuna(8), lColuna(9), lLinha(2) + 0.37, lLocal
    ImprimeCentralizado "Estoque", lColuna(9), lColuna(10), lLinha(2) + 0.2, lLocal
    ImprimeValor "0,0", lColuna(10), lColuna(11), lLinha(3) + 0.07, lLocal
    
    'Imprime linha horizontal da grade A
    Printer.Line (lColuna(0), lLinha(4))-(lColuna(4), lLinha(4)), RGB(lR, lG, lB)
    Printer.DrawWidth = 6
    Printer.Line (lColuna(4), lLinha(4))-(lColuna(10), lLinha(4)), RGB(lR, lG, lB)
    Printer.DrawWidth = 3
    Printer.Line (lColuna(10), lLinha(4))-(lColuna(11), lLinha(4)), RGB(lR, lG, lB)
    
    'Imprime 3ª linha de dados da grade A
    ImprimeTexto "Alcool Hidratado", lColuna(0), lColuna(1), lLinha(4) + 0.07, lLocal
    ImprimeValor Format(lEncerranteI(3), "#,###,##0.0"), lColuna(1), lColuna(2), lLinha(4) + 0.07, lLocal
    ImprimeValor Format(lEncerranteF(3), "#,###,##0.0"), lColuna(2), lColuna(3), lLinha(4) + 0.07, lLocal
    ImprimeValor Format(lQuantidade(3), "#,###,##0.0"), lColuna(3), lColuna(4), lLinha(4) + 0.07, lLocal
    ImprimeValor "0,0", lColuna(10), lColuna(11), lLinha(4) + 0.07, lLocal
    
    'Imprime linha horizontal da grade A
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(5))-(lColuna(4), lLinha(5)), RGB(lR, lG, lB)
    Printer.Line (lColuna(10), lLinha(5))-(lColuna(11), lLinha(5)), RGB(lR, lG, lB)
    
    'Imprime 4ª linha de dados da grade A
    ImprimeTexto "Alcool Hidratado", lColuna(0), lColuna(1), lLinha(5) + 0.07, lLocal
    ImprimeValor Format(lEncerranteI(4), "#,###,##0.0"), lColuna(1), lColuna(2), lLinha(5) + 0.07, lLocal
    ImprimeValor Format(lEncerranteF(4), "#,###,##0.0"), lColuna(2), lColuna(3), lLinha(5) + 0.07, lLocal
    ImprimeValor Format(lQuantidade(4), "#,###,##0.0"), lColuna(3), lColuna(4), lLinha(5) + 0.07, lLocal
    ImprimeValor Format(lQuantidade(3) + lQuantidade(4), "#,###,##0.0"), lColuna(4), lColuna(5), lLinha(4) + 0.37, lLocal
    ImprimeValor Format(lMedidaA, "#,###,##0.0"), lColuna(5), lColuna(6), lLinha(4) + 0.37, lLocal
    ImprimeValor "0,0", lColuna(6), lColuna(7), lLinha(4) + 0.37, lLocal
    ImprimeValor "0,0", lColuna(7), lColuna(8), lLinha(4) + 0.37, lLocal
    ImprimeTexto "R$", lColuna(8), lColuna(9), lLinha(4) + 0.37, lLocal
    ImprimeValor Format(lPrecoCusto(4), "###,##0.0000"), lColuna(8), lColuna(9), lLinha(4) + 0.37, lLocal
    ImprimeCentralizado "Estoque", lColuna(9), lColuna(10), lLinha(4) + 0.2, lLocal
    ImprimeValor "0,0", lColuna(10), lColuna(11), lLinha(5) + 0.07, lLocal
    
    'Imprime linha horizontal da grade A
    Printer.Line (lColuna(0), lLinha(6))-(lColuna(4), lLinha(6)), RGB(lR, lG, lB)
    Printer.DrawWidth = 6
    Printer.Line (lColuna(4), lLinha(6))-(lColuna(10), lLinha(6)), RGB(lR, lG, lB)
    Printer.DrawWidth = 3
    Printer.Line (lColuna(10), lLinha(6))-(lColuna(11), lLinha(6)), RGB(lR, lG, lB)
    
    'Imprime 5ª linha de dados da grade A
    ImprimeTexto "Gasolina C", lColuna(0), lColuna(1), lLinha(6) + 0.07, lLocal
    ImprimeValor Format(lEncerranteI(5), "#,###,##0.0"), lColuna(1), lColuna(2), lLinha(6) + 0.07, lLocal
    ImprimeValor Format(lEncerranteF(5), "#,###,##0.0"), lColuna(2), lColuna(3), lLinha(6) + 0.07, lLocal
    ImprimeValor Format(lQuantidade(5), "#,###,##0.0"), lColuna(3), lColuna(4), lLinha(6) + 0.07, lLocal
    ImprimeValor "0,0", lColuna(10), lColuna(11), lLinha(6) + 0.07, lLocal
    
    'Imprime linha horizontal da grade A
    Printer.Line (lColuna(0), lLinha(7))-(lColuna(4), lLinha(7)), RGB(lR, lG, lB)
    Printer.Line (lColuna(10), lLinha(7))-(lColuna(11), lLinha(7)), RGB(lR, lG, lB)
    
    'Imprime 6ª linha de dados da grade A
    ImprimeTexto "Gasolina C", lColuna(0), lColuna(1), lLinha(7) + 0.07, lLocal
    ImprimeValor Format(lEncerranteI(6), "#,###,##0.0"), lColuna(1), lColuna(2), lLinha(7) + 0.07, lLocal
    ImprimeValor Format(lEncerranteF(6), "#,###,##0.0"), lColuna(2), lColuna(3), lLinha(7) + 0.07, lLocal
    ImprimeValor Format(lQuantidade(6), "#,###,##0.0"), lColuna(3), lColuna(4), lLinha(7) + 0.07, lLocal
    ImprimeValor Format(lQuantidade(5) + lQuantidade(6), "#,###,##0.0"), lColuna(4), lColuna(5), lLinha(6) + 0.37, lLocal
    ImprimeValor Format(lMedidaG, "#,###,##0.0"), lColuna(5), lColuna(6), lLinha(6) + 0.37, lLocal
    ImprimeValor "0,0", lColuna(6), lColuna(7), lLinha(6) + 0.37, lLocal
    ImprimeValor "0,0", lColuna(7), lColuna(8), lLinha(6) + 0.37, lLocal
    ImprimeTexto "R$", lColuna(8), lColuna(9), lLinha(6) + 0.37, lLocal
    ImprimeValor Format(lPrecoCusto(6), "###,##0.0000"), lColuna(8), lColuna(9), lLinha(6) + 0.37, lLocal
    ImprimeCentralizado "Estoque", lColuna(9), lColuna(10), lLinha(6) + 0.2, lLocal
    ImprimeValor "0,0", lColuna(10), lColuna(11), lLinha(7) + 0.07, lLocal
    
    'Imprime linha horizontal da grade A
    Printer.DrawWidth = 6
    Printer.Line (lColuna(0), lLinha(8))-(lColuna(11), lLinha(8)), RGB(lR, lG, lB)
    
    'Linhas verticais (grade A)
    Printer.Line (lColuna(0), lLinha(0))-(lColuna(0), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(1), lLinha(1))-(lColuna(1), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(2), lLinha(1))-(lColuna(2), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(3), lLinha(1))-(lColuna(3), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(4), lLinha(1))-(lColuna(4), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(5), lLinha(1))-(lColuna(5), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(6), lLinha(1))-(lColuna(6), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(7), lLinha(1))-(lColuna(7), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(8), lLinha(1))-(lColuna(8), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(9), lLinha(1))-(lColuna(9), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(10), lLinha(0))-(lColuna(10), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(11), lLinha(0))-(lColuna(11), lLinha(8)), RGB(lR, lG, lB)
    
End Sub
Private Sub ImprimeGradeB()
    Dim xDataInicial As Date
    Dim i As Integer
    Dim i2 As Integer
    For i = 0 To 6
        lEncerranteI(i) = 0
        lQuantidade(i) = 0
        lPrecoCusto(i) = 0
        lPrecoVenda(i) = 0
        lValorVendaDia(i) = 0
        lValorVendaAcumulado(i) = 0
        lPrevisaoVendaQtd(i) = 0
        lPrevisaoVendaValor(i) = 0
        lReceitaBrutaEstimada(i) = 0
    Next
    
    lQtdDiasMes = 32
    Do Until IsDate(lQtdDiasMes & Mid(msk_data.Text, 3, 8))
        lQtdDiasMes = lQtdDiasMes - 1
    Loop
    
    'Monta a Select
    Set rsTabela = Nothing
    xDataInicial = "01" & Mid(msk_data.Text, 3, 8)
    lSQl = ""
    lSQl = lSQl & "SELECT Data, Periodo, [Codigo da Bomba], Abertura, Encerrante, [Quantidade da Saida], "
    lSQl = lSQl & "       [Preco de Custo], [Preco de Venda], [Tipo de Combustivel]"
    lSQl = lSQl & "  FROM Movimento_Bomba"
    lSQl = lSQl & " WHERE Empresa = " & g_empresa
    lSQl = lSQl & "   AND Data >= #" & Format(xDataInicial, "mm/dd/yyyy") & "#"
    lSQl = lSQl & "   AND Data <= #" & Format(msk_data.Text, "mm/dd/yyyy") & "#"
    lSQl = lSQl & " ORDER BY Data, Periodo, [Codigo da Bomba]"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQl)
    Do Until rsTabela.EOF
        i2 = rsTabela("Codigo da Bomba").Value
        If lEncerranteI(i2) = 0 Then
            lEncerranteI(i2) = rsTabela("Abertura").Value
        End If
        If rsTabela("Data").Value = CDate(msk_data.Text) Then
            lValorVendaDia(i2) = lValorVendaDia(i2) + rsTabela("Quantidade da Saida").Value * rsTabela("Preco de Venda").Value
            lValorVendaDia(0) = lValorVendaDia(0) + rsTabela("Quantidade da Saida").Value * rsTabela("Preco de Venda").Value
        End If
        lPrecoCusto(i2) = rsTabela("Preco de Custo").Value
        lPrecoVenda(i2) = rsTabela("Preco de Venda").Value
        lQuantidade(i2) = lQuantidade(i2) + rsTabela("Quantidade da Saida").Value
        lQuantidade(0) = lQuantidade(0) + rsTabela("Quantidade da Saida").Value
        lValorVendaAcumulado(i2) = lValorVendaAcumulado(i2) + rsTabela("Quantidade da Saida").Value * rsTabela("Preco de Venda").Value
        lValorVendaAcumulado(0) = lValorVendaAcumulado(0) + rsTabela("Quantidade da Saida").Value * rsTabela("Preco de Venda").Value
        rsTabela.MoveNext
    Loop
    
    For i = 0 To 6
        lPrevisaoVendaQtd(i) = lQuantidade(i) / Day(msk_data.Text) * lQtdDiasMes
    Next
    'Coordenada inicial da linha da grade B (em centimetros)
    lLinhaI = lLinhaI + 3.6
    
    'Coordenadas das colunas da grade B (em centimetros)
    lColuna(0) = lColunaI + 0
    lColuna(1) = lColunaI + 2.2
    lColuna(2) = lColunaI + 3.9
    lColuna(3) = lColunaI + 5.6
    lColuna(4) = lColunaI + 7.3
    lColuna(5) = lColunaI + 9.6
    lColuna(6) = lColunaI + 11.44
    lColuna(7) = lColunaI + 13.28
    lColuna(8) = lColunaI + 15.12
    lColuna(9) = lColunaI + 16.96
    lColuna(10) = lColunaI + 18.8
    lColuna(11) = lColunaI + 20.5
    
    'Coordenadas das Linhas da grade B (em centimetros)
    lLinha(0) = lLinhaI + 0
    lLinha(1) = lLinhaI + 0.62
    lLinha(2) = lLinhaI + 0.93
    lLinha(3) = lLinhaI + 1.24
    lLinha(4) = lLinhaI + 1.55
    lLinha(5) = lLinhaI + 1.86
    lLinha(6) = lLinhaI + 2.17
    lLinha(7) = lLinhaI + 2.48
    lLinha(8) = lLinhaI + 2.79
    
    'Linhas horizontais (grade B)
    Printer.DrawWidth = 6
    Printer.Line (lColuna(0), lLinha(0))-(lColuna(11), lLinha(0)), RGB(lR, lG, lB)
    
    
   
    'Imprime os dados do cabeçalho da grade B
    Printer.FontBold = False
    Printer.FontSize = 6
    ImprimeCentralizado "Produto", lColuna(0), lColuna(1), lLinha(0) + 0.21, lLocal
    ImprimeCentralizado "Iniciante Mês", lColuna(1), lColuna(2), lLinha(0) + 0.21, lLocal
    ImprimeCentralizado "Volume Saída", lColuna(2), lColuna(3), lLinha(0) + 0.1, lLocal
    ImprimeCentralizado "Mês", lColuna(2), lColuna(3), lLinha(0) + 0.36, lLocal
    ImprimeCentralizado "Vendas", lColuna(3), lColuna(4), lLinha(0) + 0.1, lLocal
    ImprimeCentralizado "Acumuladas", lColuna(3), lColuna(4), lLinha(0) + 0.36, lLocal
    ImprimeCentralizado "Previsão de Vendas", lColuna(4), lColuna(5), lLinha(0) + 0.1, lLocal
    ImprimeCentralizado "Até Final Mês", lColuna(4), lColuna(5), lLinha(0) + 0.36, lLocal
    ImprimeCentralizado "Estoque", lColuna(5), lColuna(6), lLinha(0) + 0.1, lLocal
    ImprimeCentralizado "Medição", lColuna(5), lColuna(6), lLinha(0) + 0.36, lLocal
    ImprimeCentralizado "Diferenças", lColuna(6), lColuna(7), lLinha(0) + 0.1, lLocal
    ImprimeCentralizado "Acumuladas", lColuna(6), lColuna(7), lLinha(0) + 0.36, lLocal
    ImprimeCentralizado "Compras Acum.", lColuna(7), lColuna(8), lLinha(0) + 0.21, lLocal
    ImprimeCentralizado "Valor Estoque", lColuna(8), lColuna(9), lLinha(0) + 0.21, lLocal
    ImprimeCentralizado "Média Venda", lColuna(9), lColuna(10), lLinha(0) + 0.1, lLocal
    ImprimeCentralizado "Diária", lColuna(9), lColuna(10), lLinha(0) + 0.36, lLocal
    ImprimeCentralizado "Média Venda", lColuna(10), lColuna(11), lLinha(0) + 0.1, lLocal
    ImprimeCentralizado "Percentual", lColuna(10), lColuna(11), lLinha(0) + 0.36, lLocal
    
    'Imprime 1ª linha horizontal da grade B
    Printer.Line (lColuna(0), lLinha(1))-(lColuna(11), lLinha(1)), RGB(lR, lG, lB)
    
    'Imprime 1ª linha de dados da grade B
    ImprimeTexto "Gasolina C", lColuna(0), lColuna(1), lLinha(1) + 0.04, lLocal
    ImprimeValor Format(lEncerranteI(1), "#,###,##0.0"), lColuna(1), lColuna(2), lLinha(1) + 0.04, lLocal
    ImprimeValor Format(lQuantidade(1), "#,###,##0.0"), lColuna(2), lColuna(3), lLinha(1) + 0.04, lLocal
    ImprimeValor Format(lQuantidade(1), "#,###,##0.0"), lColuna(3), lColuna(4), lLinha(1) + 0.04, lLocal
    ImprimeValor Format(lPrevisaoVendaQtd(1), "#,###,##0.0"), lColuna(4), lColuna(5), lLinha(1) + 0.04, lLocal
    
    'Imprime 2ª linha horizontal da grade B
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(2))-(lColuna(5), lLinha(2)), RGB(lR, lG, lB)
    
    'Imprime 2ª linha de dados da grade B
    ImprimeTexto "Gasolina C", lColuna(0), lColuna(1), lLinha(2) + 0.04, lLocal
    ImprimeValor Format(lEncerranteI(2), "#,###,##0.0"), lColuna(1), lColuna(2), lLinha(2) + 0.04, lLocal
    ImprimeValor Format(lQuantidade(2), "#,###,##0.0"), lColuna(2), lColuna(3), lLinha(2) + 0.04, lLocal
    ImprimeValor Format(lQuantidade(2), "#,###,##0.0"), lColuna(3), lColuna(4), lLinha(2) + 0.04, lLocal
    ImprimeValor Format(lPrevisaoVendaQtd(2), "#,###,##0.0"), lColuna(4), lColuna(5), lLinha(2) + 0.04, lLocal
    lPercentual = (lPrevisaoVendaQtd(1) + lPrevisaoVendaQtd(2)) / lQtdDiasMes
    ImprimeValor Format(lPercentual, "#,###,##0.0"), lColuna(9), lColuna(10), lLinha(2) + 0.04, lLocal
    lPercentual = (lPrevisaoVendaQtd(1) + lPrevisaoVendaQtd(2)) * 100 / lPrevisaoVendaQtd(0)
    ImprimeValor Format(lPercentual, "##0.00") & "%", lColuna(10), lColuna(11), lLinha(2) + 0.04, lLocal
    
    'Imprime 3ª linha horizontal da grade B
    Printer.Line (lColuna(0), lLinha(3))-(lColuna(5), lLinha(3)), RGB(lR, lG, lB)
    Printer.DrawWidth = 6
    Printer.Line (lColuna(5), lLinha(3))-(lColuna(11), lLinha(3)), RGB(lR, lG, lB)
    
    'Imprime 3ª linha de dados da grade B
    ImprimeTexto "Alcool Hidratado", lColuna(0), lColuna(1), lLinha(3) + 0.04, lLocal
    ImprimeValor Format(lEncerranteI(3), "#,###,##0.0"), lColuna(1), lColuna(2), lLinha(3) + 0.04, lLocal
    ImprimeValor Format(lQuantidade(3), "#,###,##0.0"), lColuna(2), lColuna(3), lLinha(3) + 0.04, lLocal
    ImprimeValor Format(lQuantidade(3), "#,###,##0.0"), lColuna(3), lColuna(4), lLinha(3) + 0.04, lLocal
    ImprimeValor Format(lPrevisaoVendaQtd(3), "#,###,##0.0"), lColuna(4), lColuna(5), lLinha(3) + 0.04, lLocal
    
    'Imprime 4ª linha horizontal da grade B
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(4))-(lColuna(5), lLinha(4)), RGB(lR, lG, lB)
    
    'Imprime 4ª linha de dados da grade B
    ImprimeTexto "Alcool Hidratado", lColuna(0), lColuna(1), lLinha(4) + 0.04, lLocal
    ImprimeValor Format(lEncerranteI(4), "#,###,##0.0"), lColuna(1), lColuna(2), lLinha(4) + 0.04, lLocal
    ImprimeValor Format(lQuantidade(4), "#,###,##0.0"), lColuna(2), lColuna(3), lLinha(4) + 0.04, lLocal
    ImprimeValor Format(lQuantidade(4), "#,###,##0.0"), lColuna(3), lColuna(4), lLinha(4) + 0.04, lLocal
    ImprimeValor Format(lPrevisaoVendaQtd(4), "#,###,##0.0"), lColuna(4), lColuna(5), lLinha(4) + 0.04, lLocal
    ImprimeValor Format(lMedidaA, "#,###,##0.0"), lColuna(5), lColuna(6), lLinha(3) + 0.37, lLocal
    ImprimeValor "0,0", lColuna(6), lColuna(7), lLinha(3) + 0.37, lLocal
    ImprimeValor "0,0", lColuna(7), lColuna(8), lLinha(3) + 0.37, lLocal
    ImprimeTexto "R$", lColuna(8), lColuna(9), lLinha(3) + 0.37, lLocal
    ImprimeValor Format(lMedidaA * lPrecoCusto(2), "#,###,##0.00"), lColuna(8), lColuna(9), lLinha(3) + 0.37, lLocal
    lPercentual = (lPrevisaoVendaQtd(3) + lPrevisaoVendaQtd(4)) / lQtdDiasMes
    ImprimeValor Format(lPercentual, "#,###,##0.0"), lColuna(9), lColuna(10), lLinha(4) + 0.04, lLocal
    lPercentual = (lPrevisaoVendaQtd(3) + lPrevisaoVendaQtd(4)) * 100 / lPrevisaoVendaQtd(0)
    ImprimeValor Format(lPercentual, "##0.00") & "%", lColuna(10), lColuna(11), lLinha(4) + 0.04, lLocal
    
    'Imprime 5ª linha horizontal da grade B
    Printer.Line (lColuna(0), lLinha(5))-(lColuna(5), lLinha(5)), RGB(lR, lG, lB)
    Printer.DrawWidth = 6
    Printer.Line (lColuna(5), lLinha(5))-(lColuna(11), lLinha(5)), RGB(lR, lG, lB)
    
    'Imprime 5ª linha de dados da grade B
    ImprimeTexto "Gasolina C", lColuna(0), lColuna(1), lLinha(5) + 0.04, lLocal
    ImprimeValor Format(lEncerranteI(5), "#,###,##0.0"), lColuna(1), lColuna(2), lLinha(5) + 0.04, lLocal
    ImprimeValor Format(lQuantidade(5), "#,###,##0.0"), lColuna(2), lColuna(3), lLinha(5) + 0.04, lLocal
    ImprimeValor Format(lQuantidade(5), "#,###,##0.0"), lColuna(3), lColuna(4), lLinha(5) + 0.04, lLocal
    ImprimeValor Format(lPrevisaoVendaQtd(5), "#,###,##0.0"), lColuna(4), lColuna(5), lLinha(5) + 0.04, lLocal
    
    'Imprime 6ª linha horizontal da grade B
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(6))-(lColuna(5), lLinha(6)), RGB(lR, lG, lB)
    
    'Imprime 6ª linha de dados da grade B
    ImprimeTexto "Gasolina C", lColuna(0), lColuna(1), lLinha(6) + 0.04, lLocal
    ImprimeValor Format(lEncerranteI(6), "#,###,##0.0"), lColuna(1), lColuna(2), lLinha(6) + 0.04, lLocal
    ImprimeValor Format(lQuantidade(6), "#,###,##0.0"), lColuna(2), lColuna(3), lLinha(6) + 0.04, lLocal
    ImprimeValor Format(lQuantidade(6), "#,###,##0.0"), lColuna(3), lColuna(4), lLinha(6) + 0.04, lLocal
    ImprimeValor Format(lPrevisaoVendaQtd(6), "#,###,##0.0"), lColuna(4), lColuna(5), lLinha(6) + 0.04, lLocal
    ImprimeValor Format(lMedidaG, "#,###,##0.0"), lColuna(5), lColuna(6), lLinha(5) + 0.37, lLocal
    ImprimeValor "0,0", lColuna(6), lColuna(7), lLinha(5) + 0.37, lLocal
    ImprimeValor "0,0", lColuna(7), lColuna(8), lLinha(5) + 0.37, lLocal
    ImprimeTexto "R$", lColuna(8), lColuna(9), lLinha(5) + 0.37, lLocal
    ImprimeValor Format(lMedidaG * lPrecoCusto(6), "#,###,##0.00"), lColuna(8), lColuna(9), lLinha(5) + 0.37, lLocal
    lPercentual = (lPrevisaoVendaQtd(5) + lPrevisaoVendaQtd(6)) / lQtdDiasMes
    ImprimeValor Format(lPercentual, "#,###,##0.0"), lColuna(9), lColuna(10), lLinha(6) + 0.04, lLocal
    lPercentual = (lPrevisaoVendaQtd(5) + lPrevisaoVendaQtd(6)) * 100 / lPrevisaoVendaQtd(0)
    ImprimeValor Format(lPercentual, "##0.00") & "%", lColuna(10), lColuna(11), lLinha(6) + 0.04, lLocal
    
    'Imprime 7ª linha horizontal da grade B
    Printer.Line (lColuna(0), lLinha(7))-(lColuna(11), lLinha(7)), RGB(lR, lG, lB)
    
    'Imprime 7ª linha de dados da grade B
    ImprimeTexto "Total", lColuna(0), lColuna(1), lLinha(7) + 0.04, lLocal
    ImprimeCentralizado "***", lColuna(1), lColuna(2), lLinha(7) + 0.04, lLocal
    ImprimeValor Format(lQuantidade(0), "#,###,##0.0"), lColuna(2), lColuna(3), lLinha(7) + 0.04, lLocal
    ImprimeValor Format(lQuantidade(0), "#,###,##0.0"), lColuna(3), lColuna(4), lLinha(7) + 0.04, lLocal
    ImprimeValor Format(lPrevisaoVendaQtd(0), "#,###,##0.0"), lColuna(4), lColuna(5), lLinha(7) + 0.04, lLocal
    ImprimeValor Format(lMedidaA + lMedidaG, "#,###,##0.0"), lColuna(5), lColuna(6), lLinha(7) + 0.04, lLocal
    ImprimeValor "0,0", lColuna(6), lColuna(7), lLinha(7) + 0.04, lLocal
    ImprimeValor "0,0", lColuna(7), lColuna(8), lLinha(7) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(8), lColuna(9), lLinha(7) + 0.04, lLocal
    ImprimeValor Format(lMedidaA * lPrecoCusto(2) + lMedidaG * lPrecoCusto(6), "#,###,##0.00"), lColuna(8), lColuna(9), lLinha(7) + 0.04, lLocal
    lPercentual = lPrevisaoVendaQtd(0) / lQtdDiasMes
    ImprimeValor Format(lPercentual, "#,###,##0.0"), lColuna(9), lColuna(10), lLinha(7) + 0.04, lLocal
    ImprimeValor "100%", lColuna(10), lColuna(11), lLinha(7) + 0.04, lLocal
    
    
    
    'Imprime 8ª linha horizontal da grade B
    Printer.DrawWidth = 6
    Printer.Line (lColuna(0), lLinha(8))-(lColuna(11), lLinha(8)), RGB(lR, lG, lB)
    
    'Imprime linhas verticais da grade B
    Printer.Line (lColuna(0), lLinha(0))-(lColuna(0), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(1), lLinha(0))-(lColuna(1), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(2), lLinha(0))-(lColuna(2), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(3), lLinha(0))-(lColuna(3), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(4), lLinha(0))-(lColuna(4), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(5), lLinha(0))-(lColuna(5), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(6), lLinha(0))-(lColuna(6), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(7), lLinha(0))-(lColuna(7), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(8), lLinha(0))-(lColuna(8), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(9), lLinha(0))-(lColuna(9), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(10), lLinha(0))-(lColuna(10), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(11), lLinha(0))-(lColuna(11), lLinha(8)), RGB(lR, lG, lB)
    
End Sub
Private Sub ImprimeGradeC()
    Dim i As Integer
    'Coordenada inicial da linha da grade C (em centimetros)
    lLinhaI = lLinhaI + 2.9
    
    'Coordenadas das colunas da grade C (em centimetros)
    lColuna(0) = lColunaI + 0
    lColuna(1) = lColunaI + 2.2
    lColuna(2) = lColunaI + 3.9
    lColuna(3) = lColunaI + 5.6
    lColuna(4) = lColunaI + 7.3
    lColuna(5) = lColunaI + 9.6
    lColuna(6) = lColunaI + 11.44
    lColuna(7) = lColunaI + 13.28
    lColuna(8) = lColunaI + 15.12
    lColuna(9) = lColunaI + 16.96
    lColuna(10) = lColunaI + 18.8
    lColuna(11) = lColunaI + 20.5
    
    'Coordenadas das Linhas da grade C (em centimetros)
    lLinha(0) = lLinhaI + 0
    lLinha(1) = lLinhaI + 0.62
    lLinha(2) = lLinhaI + 0.93
    lLinha(3) = lLinhaI + 1.24
    lLinha(4) = lLinhaI + 1.55
    lLinha(5) = lLinhaI + 1.86
    lLinha(6) = lLinhaI + 2.17
    lLinha(7) = lLinhaI + 2.48
    lLinha(8) = lLinhaI + 2.79
    
    lReceitaBrutaEstimada(0) = 0
    For i = 0 To 6
        lPrevisaoVendaValor(i) = lValorVendaAcumulado(i) / Day(msk_data.Text) * lQtdDiasMes
        lReceitaBrutaEstimada(i) = lPrevisaoVendaQtd(i) * (lPrecoVenda(i) - lPrecoCusto(i))
    Next
    For i = 1 To 6
        lReceitaBrutaEstimada(0) = lReceitaBrutaEstimada(0) + Format(lReceitaBrutaEstimada(i), "00000000.00")
    Next
    
    Printer.DrawWidth = 6
    Printer.Line (lColuna(0), lLinha(0))-(lColuna(11), lLinha(0)), RGB(lR, lG, lB)
    
    'Imprime os dados do cabeçalho da grade C
    Printer.FontBold = False
    Printer.FontSize = 6
    ImprimeCentralizado "Produto", lColuna(0), lColuna(1), lLinha(0) + 0.21, lLocal
    ImprimeCentralizado "Preço", lColuna(1), lColuna(2), lLinha(0) + 0.06, lLocal
    ImprimeCentralizado "Bomba", lColuna(1), lColuna(2), lLinha(0) + 0.36, lLocal
    ImprimeCentralizado "Vendas", lColuna(2), lColuna(3), lLinha(0) + 0.06, lLocal
    ImprimeCentralizado "Financeiras", lColuna(2), lColuna(3), lLinha(0) + 0.36, lLocal
    ImprimeCentralizado "Vendas", lColuna(3), lColuna(4), lLinha(0) + 0.06, lLocal
    ImprimeCentralizado "Financ. Acum.", lColuna(3), lColuna(4), lLinha(0) + 0.36, lLocal
    ImprimeCentralizado "Margem Bruto", lColuna(4), lColuna(5), lLinha(0) + 0.21, lLocal
    ImprimeCentralizado "Diferenças", lColuna(5), lColuna(6), lLinha(0) + 0.06, lLocal
    ImprimeCentralizado "Medição", lColuna(5), lColuna(6), lLinha(0) + 0.36, lLocal
    ImprimeCentralizado "Preço Médio", lColuna(6), lColuna(7), lLinha(0) + 0.06, lLocal
    ImprimeCentralizado "Venda", lColuna(6), lColuna(7), lLinha(0) + 0.36, lLocal
    ImprimeCentralizado "Valores Acum. s/", lColuna(7), lColuna(8), lLinha(0) + 0.06, lLocal
    ImprimeCentralizado "Compras", lColuna(7), lColuna(8), lLinha(0) + 0.36, lLocal
    ImprimeCentralizado "Previsão", lColuna(8), lColuna(9), lLinha(0) + 0.06, lLocal
    ImprimeCentralizado "Faturamento", lColuna(8), lColuna(9), lLinha(0) + 0.36, lLocal
    ImprimeCentralizado "Preço Compra", lColuna(9), lColuna(10), lLinha(0) + 0.06, lLocal
    ImprimeCentralizado "Médio", lColuna(9), lColuna(10), lLinha(0) + 0.36, lLocal
    ImprimeCentralizado "Receita Bruta", lColuna(10), lColuna(11), lLinha(0) + 0.06, lLocal
    ImprimeCentralizado "Estimada", lColuna(10), lColuna(11), lLinha(0) + 0.36, lLocal
    
    'Imprime linha horizontal da grade C
    Printer.Line (lColuna(0), lLinha(1))-(lColuna(11), lLinha(1)), RGB(lR, lG, lB)
    
    'Imprime 1ª linha de dados da grade C
    ImprimeTexto "Gasolina C", lColuna(0), lColuna(1), lLinha(1) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(1), lColuna(2), lLinha(1) + 0.04, lLocal
    ImprimeValor Format(lPrecoVenda(1), "#,###,##0.00"), lColuna(1), lColuna(2), lLinha(1) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(1) + 0.04, lLocal
    ImprimeValor Format(lValorVendaDia(1), "#,###,##0.00"), lColuna(2), lColuna(3), lLinha(1) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(3), lColuna(4), lLinha(1) + 0.04, lLocal
    ImprimeValor Format(lValorVendaAcumulado(1), "#,###,##0.00"), lColuna(3), lColuna(4), lLinha(1) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(1) + 0.04, lLocal
    ImprimeValor Format(lValorVendaAcumulado(1) / lQuantidade(1), "#,###,##0.0000"), lColuna(6), lColuna(7), lLinha(1) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(8), lColuna(9), lLinha(1) + 0.04, lLocal
    ImprimeValor Format(lPrevisaoVendaValor(1), "#,###,##0.00"), lColuna(8), lColuna(9), lLinha(1) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(10), lColuna(11), lLinha(1) + 0.04, lLocal
    ImprimeValor Format(lReceitaBrutaEstimada(1), "#,###,##0.00"), lColuna(10), lColuna(11), lLinha(1) + 0.04, lLocal
    
    'Imprime linha horizontal da grade C
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(2))-(lColuna(4), lLinha(2)), RGB(lR, lG, lB)
    Printer.Line (lColuna(6), lLinha(2))-(lColuna(7), lLinha(2)), RGB(lR, lG, lB)
    Printer.Line (lColuna(8), lLinha(2))-(lColuna(9), lLinha(2)), RGB(lR, lG, lB)
    Printer.Line (lColuna(10), lLinha(2))-(lColuna(11), lLinha(2)), RGB(lR, lG, lB)
    
    'Imprime 2ª linha de dados da grade C
    ImprimeTexto "Gasolina C", lColuna(0), lColuna(1), lLinha(2) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(1), lColuna(2), lLinha(2) + 0.04, lLocal
    ImprimeValor Format(lPrecoVenda(2), "#,###,##0.00"), lColuna(1), lColuna(2), lLinha(2) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(2) + 0.04, lLocal
    ImprimeValor Format(lValorVendaDia(2), "#,###,##0.00"), lColuna(2), lColuna(3), lLinha(2) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(3), lColuna(4), lLinha(2) + 0.04, lLocal
    ImprimeValor Format(lValorVendaAcumulado(2), "#,###,##0.00"), lColuna(3), lColuna(4), lLinha(2) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(4), lColuna(5), lLinha(1) + 0.37, lLocal
    ImprimeValor Format(lPrecoVenda(2) - lPrecoCusto(2), "#,###,##0.00"), lColuna(4), lColuna(5), lLinha(1) + 0.37, lLocal
    ImprimeTexto "R$", lColuna(5), lColuna(6), lLinha(1) + 0.37, lLocal
    ImprimeValor "000,0", lColuna(5), lColuna(6), lLinha(1) + 0.37, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(2) + 0.04, lLocal
    ImprimeValor Format(lValorVendaAcumulado(2) / lQuantidade(2), "#,###,##0.0000"), lColuna(6), lColuna(7), lLinha(2) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(7), lColuna(8), lLinha(1) + 0.37, lLocal
    ImprimeValor "000,0", lColuna(7), lColuna(8), lLinha(1) + 0.37, lLocal
    ImprimeTexto "R$", lColuna(8), lColuna(9), lLinha(2) + 0.04, lLocal
    ImprimeValor Format(lPrevisaoVendaValor(2), "#,###,##0.00"), lColuna(8), lColuna(9), lLinha(2) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(9), lColuna(10), lLinha(1) + 0.37, lLocal
    ImprimeValor "000,0", lColuna(9), lColuna(10), lLinha(1) + 0.37, lLocal
    ImprimeTexto "R$", lColuna(10), lColuna(11), lLinha(2) + 0.04, lLocal
    ImprimeValor Format(lReceitaBrutaEstimada(2), "#,###,##0.00"), lColuna(10), lColuna(11), lLinha(2) + 0.04, lLocal
    
    'Imprime linha horizontal da grade C
    Printer.Line (lColuna(0), lLinha(3))-(lColuna(4), lLinha(3)), RGB(lR, lG, lB)
    Printer.DrawWidth = 6
    Printer.Line (lColuna(4), lLinha(3))-(lColuna(10), lLinha(3)), RGB(lR, lG, lB)
    Printer.DrawWidth = 3
    Printer.Line (lColuna(10), lLinha(3))-(lColuna(11), lLinha(3)), RGB(lR, lG, lB)
    
    'Imprime 3ª linha de dados da grade C
    ImprimeTexto "Alcool Hidratado", lColuna(0), lColuna(1), lLinha(3) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(1), lColuna(2), lLinha(3) + 0.04, lLocal
    ImprimeValor Format(lPrecoVenda(3), "#,###,##0.00"), lColuna(1), lColuna(2), lLinha(3) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(3) + 0.04, lLocal
    ImprimeValor Format(lValorVendaDia(3), "#,###,##0.00"), lColuna(2), lColuna(3), lLinha(3) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(3), lColuna(4), lLinha(3) + 0.04, lLocal
    ImprimeValor Format(lValorVendaAcumulado(3), "#,###,##0.00"), lColuna(3), lColuna(4), lLinha(3) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(3) + 0.04, lLocal
    ImprimeValor Format(lValorVendaAcumulado(3) / lQuantidade(3), "#,###,##0.0000"), lColuna(6), lColuna(7), lLinha(3) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(8), lColuna(9), lLinha(3) + 0.04, lLocal
    ImprimeValor Format(lPrevisaoVendaValor(3), "#,###,##0.00"), lColuna(8), lColuna(9), lLinha(3) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(10), lColuna(11), lLinha(3) + 0.04, lLocal
    ImprimeValor Format(lReceitaBrutaEstimada(3), "#,###,##0.00"), lColuna(10), lColuna(11), lLinha(3) + 0.04, lLocal
    
    'Imprime linha horizontal da grade C
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(4))-(lColuna(4), lLinha(4)), RGB(lR, lG, lB)
    Printer.Line (lColuna(6), lLinha(4))-(lColuna(7), lLinha(4)), RGB(lR, lG, lB)
    Printer.Line (lColuna(8), lLinha(4))-(lColuna(9), lLinha(4)), RGB(lR, lG, lB)
    Printer.Line (lColuna(10), lLinha(4))-(lColuna(11), lLinha(4)), RGB(lR, lG, lB)
    
    'Imprime 4ª linha de dados da grade C
    ImprimeTexto "Alcool Hidratado", lColuna(0), lColuna(1), lLinha(4) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(1), lColuna(2), lLinha(4) + 0.04, lLocal
    ImprimeValor Format(lPrecoVenda(4), "#,###,##0.00"), lColuna(1), lColuna(2), lLinha(4) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(4) + 0.04, lLocal
    ImprimeValor Format(lValorVendaDia(4), "#,###,##0.00"), lColuna(2), lColuna(3), lLinha(4) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(3), lColuna(4), lLinha(4) + 0.04, lLocal
    ImprimeValor Format(lValorVendaAcumulado(4), "#,###,##0.00"), lColuna(3), lColuna(4), lLinha(4) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(4), lColuna(5), lLinha(3) + 0.37, lLocal
    ImprimeValor Format(lPrecoVenda(4) - lPrecoCusto(4), "#,###,##0.00"), lColuna(4), lColuna(5), lLinha(3) + 0.37, lLocal
    ImprimeTexto "R$", lColuna(5), lColuna(6), lLinha(3) + 0.37, lLocal
    ImprimeValor "000,0", lColuna(5), lColuna(6), lLinha(3) + 0.37, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(4) + 0.04, lLocal
    ImprimeValor Format(lValorVendaAcumulado(4) / lQuantidade(4), "#,###,##0.0000"), lColuna(6), lColuna(7), lLinha(4) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(7), lColuna(8), lLinha(3) + 0.37, lLocal
    ImprimeValor "000,0", lColuna(7), lColuna(8), lLinha(3) + 0.37, lLocal
    ImprimeTexto "R$", lColuna(8), lColuna(9), lLinha(4) + 0.04, lLocal
    ImprimeValor Format(lPrevisaoVendaValor(4), "#,###,##0.00"), lColuna(8), lColuna(9), lLinha(4) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(9), lColuna(10), lLinha(3) + 0.37, lLocal
    ImprimeValor "000,0", lColuna(9), lColuna(10), lLinha(3) + 0.37, lLocal
    ImprimeTexto "R$", lColuna(10), lColuna(11), lLinha(4) + 0.04, lLocal
    ImprimeValor Format(lReceitaBrutaEstimada(4), "#,###,##0.00"), lColuna(10), lColuna(11), lLinha(4) + 0.04, lLocal
    
    'Imprime linha horizontal da grade C
    Printer.Line (lColuna(0), lLinha(5))-(lColuna(4), lLinha(5)), RGB(lR, lG, lB)
    Printer.DrawWidth = 6
    Printer.Line (lColuna(4), lLinha(5))-(lColuna(10), lLinha(5)), RGB(lR, lG, lB)
    Printer.DrawWidth = 3
    Printer.Line (lColuna(10), lLinha(5))-(lColuna(11), lLinha(5)), RGB(lR, lG, lB)
    
    'Imprime 5ª linha de dados da grade C
    ImprimeTexto "Gasolina C", lColuna(0), lColuna(1), lLinha(5) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(1), lColuna(2), lLinha(5) + 0.04, lLocal
    ImprimeValor Format(lPrecoVenda(5), "#,###,##0.00"), lColuna(1), lColuna(2), lLinha(5) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(5) + 0.04, lLocal
    ImprimeValor Format(lValorVendaDia(5), "#,###,##0.00"), lColuna(2), lColuna(3), lLinha(5) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(3), lColuna(4), lLinha(5) + 0.04, lLocal
    ImprimeValor Format(lValorVendaAcumulado(5), "#,###,##0.00"), lColuna(3), lColuna(4), lLinha(5) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(5) + 0.04, lLocal
    ImprimeValor Format(lValorVendaAcumulado(5) / lQuantidade(5), "#,###,##0.0000"), lColuna(6), lColuna(7), lLinha(5) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(8), lColuna(9), lLinha(5) + 0.04, lLocal
    ImprimeValor Format(lPrevisaoVendaValor(5), "#,###,##0.00"), lColuna(8), lColuna(9), lLinha(5) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(10), lColuna(11), lLinha(5) + 0.04, lLocal
    ImprimeValor Format(lReceitaBrutaEstimada(5), "#,###,##0.00"), lColuna(10), lColuna(11), lLinha(5) + 0.04, lLocal
    
    'Imprime linha horizontal da grade C
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(6))-(lColuna(4), lLinha(6)), RGB(lR, lG, lB)
    Printer.Line (lColuna(6), lLinha(6))-(lColuna(7), lLinha(6)), RGB(lR, lG, lB)
    Printer.Line (lColuna(8), lLinha(6))-(lColuna(9), lLinha(6)), RGB(lR, lG, lB)
    Printer.Line (lColuna(10), lLinha(6))-(lColuna(11), lLinha(6)), RGB(lR, lG, lB)
    
    'Imprime 6ª linha de dados da grade C
    ImprimeTexto "Gasolina C", lColuna(0), lColuna(1), lLinha(6) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(1), lColuna(2), lLinha(6) + 0.04, lLocal
    ImprimeValor Format(lPrecoVenda(6), "#,###,##0.00"), lColuna(1), lColuna(2), lLinha(6) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(6) + 0.04, lLocal
    ImprimeValor Format(lValorVendaDia(6), "#,###,##0.00"), lColuna(2), lColuna(3), lLinha(6) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(3), lColuna(4), lLinha(6) + 0.04, lLocal
    ImprimeValor Format(lValorVendaAcumulado(6), "#,###,##0.00"), lColuna(3), lColuna(4), lLinha(6) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(4), lColuna(5), lLinha(5) + 0.37, lLocal
    ImprimeValor Format(lPrecoVenda(6) - lPrecoCusto(6), "#,###,##0.00"), lColuna(4), lColuna(5), lLinha(5) + 0.37, lLocal
    ImprimeTexto "R$", lColuna(5), lColuna(6), lLinha(5) + 0.37, lLocal
    ImprimeValor "000,0", lColuna(5), lColuna(6), lLinha(5) + 0.37, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(6) + 0.04, lLocal
    ImprimeValor Format(lValorVendaAcumulado(6) / lQuantidade(6), "#,###,##0.0000"), lColuna(6), lColuna(7), lLinha(6) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(7), lColuna(8), lLinha(5) + 0.37, lLocal
    ImprimeValor "000,0", lColuna(7), lColuna(8), lLinha(5) + 0.37, lLocal
    ImprimeTexto "R$", lColuna(8), lColuna(9), lLinha(6) + 0.04, lLocal
    ImprimeValor Format(lPrevisaoVendaValor(6), "#,###,##0.00"), lColuna(8), lColuna(9), lLinha(6) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(9), lColuna(10), lLinha(5) + 0.37, lLocal
    ImprimeValor "000,0", lColuna(9), lColuna(10), lLinha(5) + 0.37, lLocal
    ImprimeTexto "R$", lColuna(10), lColuna(11), lLinha(6) + 0.04, lLocal
    ImprimeValor Format(lReceitaBrutaEstimada(6), "#,###,##0.00"), lColuna(10), lColuna(11), lLinha(6) + 0.04, lLocal
    
    'Imprime linha horizontal da grade C
    Printer.Line (lColuna(0), lLinha(7))-(lColuna(4), lLinha(7)), RGB(lR, lG, lB)
    Printer.DrawWidth = 6
    Printer.Line (lColuna(4), lLinha(7))-(lColuna(10), lLinha(7)), RGB(lR, lG, lB)
    Printer.DrawWidth = 3
    Printer.Line (lColuna(10), lLinha(7))-(lColuna(11), lLinha(7)), RGB(lR, lG, lB)
    
    'Imprime 7ª linha de dados da grade C
    ImprimeTexto "Total", lColuna(0), lColuna(1), lLinha(7) + 0.04, lLocal
    ImprimeCentralizado "***", lColuna(1), lColuna(2), lLinha(7) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(7) + 0.04, lLocal
    ImprimeValor Format(lValorVendaDia(0), "#,###,##0.00"), lColuna(2), lColuna(3), lLinha(7) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(3), lColuna(4), lLinha(7) + 0.04, lLocal
    ImprimeValor Format(lValorVendaAcumulado(0), "#,###,##0.00"), lColuna(3), lColuna(4), lLinha(7) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(4), lColuna(5), lLinha(7) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(4), lColuna(5), lLinha(7) + 0.04, lLocal
    ImprimeValor Format(lValorVendaAcumulado(0) / lQuantidade(0), "#,###,##0.0000"), lColuna(6), lColuna(7), lLinha(7) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(7), lColuna(8), lLinha(7) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(7), lColuna(8), lLinha(7) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(8), lColuna(9), lLinha(7) + 0.04, lLocal
    ImprimeValor Format(lPrevisaoVendaValor(0), "#,###,##0.00"), lColuna(8), lColuna(9), lLinha(7) + 0.04, lLocal
    ImprimeCentralizado "***", lColuna(9), lColuna(10), lLinha(7) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(10), lColuna(11), lLinha(7) + 0.04, lLocal
    ImprimeValor Format(lReceitaBrutaEstimada(0), "#,###,##0.00"), lColuna(10), lColuna(11), lLinha(7) + 0.04, lLocal
    
    'Imprime linha horizontal da grade C
    Printer.DrawWidth = 6
    Printer.Line (lColuna(0), lLinha(8))-(lColuna(11), lLinha(8)), RGB(lR, lG, lB)
    
    'Imprime linhas verticais da grade C
    Printer.Line (lColuna(0), lLinha(0))-(lColuna(0), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(1), lLinha(0))-(lColuna(1), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(2), lLinha(0))-(lColuna(2), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(3), lLinha(0))-(lColuna(3), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(4), lLinha(0))-(lColuna(4), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(5), lLinha(0))-(lColuna(5), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(6), lLinha(0))-(lColuna(6), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(7), lLinha(0))-(lColuna(7), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(8), lLinha(0))-(lColuna(8), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(9), lLinha(0))-(lColuna(9), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(10), lLinha(0))-(lColuna(10), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(11), lLinha(0))-(lColuna(11), lLinha(8)), RGB(lR, lG, lB)
    
End Sub
Private Sub ImprimeGradeD()
  
    'Coordenada inicial da linha da grade D (em centimetros)
    lLinhaI = lLinhaI + 2.9
    
    'Coordenadas das colunas da grade D (em centimetros)
    lColuna(0) = lColunaI + 0
    lColuna(1) = lColunaI + 2.2
    lColuna(2) = lColunaI + 3.9
    lColuna(3) = lColunaI + 5.6
    lColuna(4) = lColunaI + 7.3
    lColuna(5) = lColunaI + 9.6
    lColuna(6) = lColunaI + 11.44
    lColuna(7) = lColunaI + 13.28
    lColuna(8) = lColunaI + 15.12
    lColuna(9) = lColunaI + 16.96
    lColuna(10) = lColunaI + 18.8
    lColuna(11) = lColunaI + 20.5
    
    'Coordenadas das Linhas da grade D (em centimetros)
    lLinha(0) = lLinhaI + 0
    lLinha(1) = lLinhaI + 0.31
    lLinha(2) = lLinhaI + 0.62
    lLinha(3) = lLinhaI + 0.93
    lLinha(4) = lLinhaI + 1.24
    lLinha(5) = lLinhaI + 1.55
    lLinha(6) = lLinhaI + 1.86
    lLinha(7) = lLinhaI + 2.17
    lLinha(8) = lLinhaI + 2.48
    lLinha(9) = lLinhaI + 2.79
    lLinha(10) = lLinhaI + 3.1
    lLinha(11) = lLinhaI + 3.41
    lLinha(12) = lLinhaI + 3.72
    
    Printer.DrawWidth = 6
    Printer.Line (lColuna(0), lLinha(0))-(lColuna(11), lLinha(0)), RGB(lR, lG, lB)
    
    'Imprime os dados do cabeçalho da grade D
    Printer.FontBold = False
    Printer.FontSize = 6
    ImprimeCentralizado "Contas a Receber", lColuna(0), lColuna(2), lLinha(0) + 0.04, lLocal
    ImprimeCentralizado "Acumulado", lColuna(2), lColuna(3), lLinha(0) + 0.21, lLocal
    ImprimeCentralizado "Represent. %", lColuna(3), lColuna(4), lLinha(0) + 0.21, lLocal
    ImprimeCentralizado "Contas a Pagar", lColuna(4), lColuna(6), lLinha(0) + 0.04, lLocal
    ImprimeCentralizado "Acumulado", lColuna(6), lColuna(7), lLinha(0) + 0.21, lLocal
    ImprimeCentralizado "Represent. %", lColuna(7), lColuna(8), lLinha(0) + 0.21, lLocal
    ImprimeCentralizado "Resultados", lColuna(8), lColuna(9), lLinha(0) + 0.2, lLocal
    ImprimeCentralizado "Valores", lColuna(9), lColuna(11), lLinha(0) + 0.04, lLocal
    
    'Imprime linha horizontal da grade D
    Printer.Line (lColuna(0), lLinha(1))-(lColuna(2), lLinha(1)), RGB(lR, lG, lB)
    Printer.Line (lColuna(4), lLinha(1))-(lColuna(6), lLinha(1)), RGB(lR, lG, lB)
    Printer.Line (lColuna(9), lLinha(1))-(lColuna(11), lLinha(1)), RGB(lR, lG, lB)
    
    ImprimeCentralizado "Créditos", lColuna(0), lColuna(2), lLinha(1) + 0.04, lLocal
    ImprimeCentralizado "Débitos", lColuna(4), lColuna(6), lLinha(1) + 0.04, lLocal
    ImprimeCentralizado "Estimados", lColuna(8), lColuna(9), lLinha(1) + 0.2, lLocal
    ImprimeCentralizado "Mês", lColuna(10), lColuna(11), lLinha(1) + 0.04, lLocal
    
    'Imprime linha horizontal da grade D
    Printer.Line (lColuna(0), lLinha(2))-(lColuna(8), lLinha(2)), RGB(lR, lG, lB)
    Printer.DrawWidth = 3
    Printer.Line (lColuna(10), lLinha(2))-(lColuna(11), lLinha(2)), RGB(lR, lG, lB)
    
    'Imprime 2ª linha de dados da grade D
    ImprimeTexto "Cheque a Vista", lColuna(0), lColuna(1), lLinha(2) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(1), lColuna(2), lLinha(2) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(1), lColuna(2), lLinha(2) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(2) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(2) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(3), lColuna(4), lLinha(2) + 0.04, lLocal
    ImprimeTexto "Salários/Encargos/Benef.", lColuna(4), lColuna(5), lLinha(2) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(5), lColuna(6), lLinha(2) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(5), lColuna(6), lLinha(2) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(2) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(6), lColuna(7), lLinha(2) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(7), lColuna(8), lLinha(2) + 0.04, lLocal
    ImprimeCentralizado "Mês", lColuna(9), lColuna(10), lLinha(1) + 0.21, lLocal
    ImprimeCentralizado "Acumulado", lColuna(10), lColuna(11), lLinha(2) + 0.04, lLocal
    
    'Imprime linha horizontal da grade D
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(3))-(lColuna(11), lLinha(3)), RGB(lR, lG, lB)
    
    'Imprime 3ª linha de dados da grade D
    ImprimeTexto "Cheque a Prazo", lColuna(0), lColuna(1), lLinha(3) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(1), lColuna(2), lLinha(3) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(1), lColuna(2), lLinha(3) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(3) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(3) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(3), lColuna(4), lLinha(3) + 0.04, lLocal
    ImprimeTexto "Material de Limpeza", lColuna(4), lColuna(5), lLinha(3) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(5), lColuna(6), lLinha(3) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(5), lColuna(6), lLinha(3) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(3) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(6), lColuna(7), lLinha(3) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(7), lColuna(8), lLinha(3) + 0.04, lLocal
    ImprimeTexto "Lucro Líquido", lColuna(8), lColuna(9), lLinha(3) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(9), lColuna(10), lLinha(3) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(9), lColuna(10), lLinha(3) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(10), lColuna(11), lLinha(3) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(10), lColuna(11), lLinha(3) + 0.04, lLocal
    
    'Imprime linha horizontal da grade D
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(4))-(lColuna(11), lLinha(4)), RGB(lR, lG, lB)
    
    'Imprime 4ª linha de dados da grade D
    ImprimeTexto "Cartão Visa", lColuna(0), lColuna(1), lLinha(4) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(1), lColuna(2), lLinha(4) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(1), lColuna(2), lLinha(4) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(4) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(4) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(3), lColuna(4), lLinha(4) + 0.04, lLocal
    ImprimeTexto "Telegoiás", lColuna(4), lColuna(5), lLinha(4) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(5), lColuna(6), lLinha(4) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(5), lColuna(6), lLinha(4) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(4) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(6), lColuna(7), lLinha(4) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(7), lColuna(8), lLinha(4) + 0.04, lLocal
    ImprimeTexto "Investimento", lColuna(8), lColuna(9), lLinha(4) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(9), lColuna(10), lLinha(4) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(9), lColuna(10), lLinha(4) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(10), lColuna(11), lLinha(4) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(10), lColuna(11), lLinha(4) + 0.04, lLocal
    
    'Imprime linha horizontal da grade D
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(5))-(lColuna(11), lLinha(5)), RGB(lR, lG, lB)
    
    'Imprime 5ª linha de dados da grade D
    ImprimeTexto "Cartão Credicard", lColuna(0), lColuna(1), lLinha(5) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(1), lColuna(2), lLinha(5) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(1), lColuna(2), lLinha(5) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(5) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(5) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(3), lColuna(4), lLinha(5) + 0.04, lLocal
    ImprimeTexto "Seguro Real/Santhiago", lColuna(4), lColuna(5), lLinha(5) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(5), lColuna(6), lLinha(5) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(5), lColuna(6), lLinha(5) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(5) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(6), lColuna(7), lLinha(5) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(7), lColuna(8), lLinha(5) + 0.04, lLocal
    ImprimeTexto "Premium", lColuna(8), lColuna(9), lLinha(5) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(9), lColuna(10), lLinha(5) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(9), lColuna(10), lLinha(5) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(10), lColuna(11), lLinha(5) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(10), lColuna(11), lLinha(5) + 0.04, lLocal
    
    'Imprime linha horizontal da grade D
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(6))-(lColuna(11), lLinha(6)), RGB(lR, lG, lB)
    
    'Imprime 6ª linha de dados da grade D
    ImprimeTexto "Dinheiro", lColuna(0), lColuna(1), lLinha(6) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(1), lColuna(2), lLinha(6) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(1), lColuna(2), lLinha(6) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(6) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(6) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(3), lColuna(4), lLinha(6) + 0.04, lLocal
    ImprimeTexto "Saneago/Celg", lColuna(4), lColuna(5), lLinha(6) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(5), lColuna(6), lLinha(6) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(5), lColuna(6), lLinha(6) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(6) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(6), lColuna(7), lLinha(6) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(7), lColuna(8), lLinha(6) + 0.04, lLocal
    
    'Imprime linha horizontal da grade D
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(7))-(lColuna(11), lLinha(7)), RGB(lR, lG, lB)
    
    'Imprime 7ª linha de dados da grade D
    ImprimeTexto "Vales/Tikts", lColuna(0), lColuna(1), lLinha(7) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(1), lColuna(2), lLinha(7) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(1), lColuna(2), lLinha(7) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(7) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(7) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(3), lColuna(4), lLinha(7) + 0.04, lLocal
    ImprimeTexto "Material de Bombas", lColuna(4), lColuna(5), lLinha(7) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(5), lColuna(6), lLinha(7) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(5), lColuna(6), lLinha(7) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(7) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(6), lColuna(7), lLinha(7) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(7), lColuna(8), lLinha(7) + 0.04, lLocal
    
    'Imprime linha horizontal da grade D
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(8))-(lColuna(11), lLinha(8)), RGB(lR, lG, lB)
    
    'Imprime 8ª linha de dados da grade D
    ImprimeTexto "Notas", lColuna(0), lColuna(1), lLinha(8) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(1), lColuna(2), lLinha(8) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(1), lColuna(2), lLinha(8) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(8) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(8) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(3), lColuna(4), lLinha(8) + 0.04, lLocal
    ImprimeTexto "Outros", lColuna(4), lColuna(5), lLinha(8) + 0.37, lLocal
    ImprimeTexto "R$", lColuna(5), lColuna(6), lLinha(8) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(5), lColuna(6), lLinha(8) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(8) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(6), lColuna(7), lLinha(8) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(7), lColuna(8), lLinha(8) + 0.04, lLocal
    
    'Imprime linha horizontal da grade D
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(9))-(lColuna(11), lLinha(9)), RGB(lR, lG, lB)
    
    'Imprime 9ª linha de dados da grade D
    ImprimeTexto "Vales a Receber", lColuna(0), lColuna(1), lLinha(9) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(1), lColuna(2), lLinha(9) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(1), lColuna(2), lLinha(9) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(9) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(9) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(3), lColuna(4), lLinha(9) + 0.04, lLocal
    ImprimeTexto "Compra de Produtos", lColuna(4), lColuna(5), lLinha(9) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(5), lColuna(6), lLinha(9) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(5), lColuna(6), lLinha(9) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(9) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(6), lColuna(7), lLinha(9) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(7), lColuna(8), lLinha(9) + 0.04, lLocal
    
    'Imprime linha horizontal da grade D
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(10))-(lColuna(11), lLinha(10)), RGB(lR, lG, lB)
    
    'Imprime 10ª linha de dados da grade D
    ImprimeTexto "Outros", lColuna(0), lColuna(1), lLinha(10) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(1), lColuna(2), lLinha(10) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(1), lColuna(2), lLinha(10) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(10) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(10) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(3), lColuna(4), lLinha(10) + 0.04, lLocal
    ImprimeTexto "Serviços Tercerizados", lColuna(4), lColuna(5), lLinha(10) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(5), lColuna(6), lLinha(10) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(5), lColuna(6), lLinha(10) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(10) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(6), lColuna(7), lLinha(10) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(7), lColuna(8), lLinha(10) + 0.04, lLocal
  
    'Imprime linha horizontal da grade D
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(11))-(lColuna(11), lLinha(11)), RGB(lR, lG, lB)
    
    'Imprime 11ª linha de dados da grade D
    ImprimeTexto "Total", lColuna(0), lColuna(1), lLinha(11) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(1), lColuna(2), lLinha(11) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(1), lColuna(2), lLinha(11) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(11) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(11) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(3), lColuna(4), lLinha(11) + 0.04, lLocal
    ImprimeTexto "Total", lColuna(4), lColuna(5), lLinha(11) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(5), lColuna(6), lLinha(11) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(5), lColuna(6), lLinha(11) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(11) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(6), lColuna(7), lLinha(11) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(7), lColuna(8), lLinha(11) + 0.04, lLocal
    
    'Imprime linha horizontal da grade D
    Printer.DrawWidth = 6
    Printer.Line (lColuna(0), lLinha(12))-(lColuna(11), lLinha(12)), RGB(lR, lG, lB)
    
    'Imprime linhas verticais da grade D
    Printer.Line (lColuna(0), lLinha(0))-(lColuna(0), lLinha(12)), RGB(lR, lG, lB)
    Printer.Line (lColuna(1), lLinha(2))-(lColuna(1), lLinha(12)), RGB(lR, lG, lB)
    Printer.Line (lColuna(2), lLinha(0))-(lColuna(2), lLinha(12)), RGB(lR, lG, lB)
    Printer.Line (lColuna(3), lLinha(0))-(lColuna(3), lLinha(12)), RGB(lR, lG, lB)
    Printer.Line (lColuna(4), lLinha(0))-(lColuna(4), lLinha(12)), RGB(lR, lG, lB)
    Printer.Line (lColuna(5), lLinha(2))-(lColuna(5), lLinha(12)), RGB(lR, lG, lB)
    Printer.Line (lColuna(6), lLinha(0))-(lColuna(6), lLinha(12)), RGB(lR, lG, lB)
    Printer.Line (lColuna(7), lLinha(0))-(lColuna(7), lLinha(12)), RGB(lR, lG, lB)
    Printer.Line (lColuna(8), lLinha(0))-(lColuna(8), lLinha(12)), RGB(lR, lG, lB)
    Printer.Line (lColuna(9), lLinha(0))-(lColuna(9), lLinha(12)), RGB(lR, lG, lB)
    Printer.Line (lColuna(10), lLinha(1))-(lColuna(10), lLinha(12)), RGB(lR, lG, lB)
    Printer.Line (lColuna(11), lLinha(0))-(lColuna(11), lLinha(12)), RGB(lR, lG, lB)
    
End Sub
Private Sub ImprimeGradeE()
  
    'Coordenada inicial da linha da grade E (em centimetros)
    lLinhaI = lLinhaI + 3.85
    
    'Coordenadas das colunas da grade E (em centimetros)
    lColuna(0) = lColunaI + 0
    lColuna(1) = lColunaI + 2.2
    lColuna(2) = lColunaI + 3.9
    lColuna(3) = lColunaI + 5.6
    lColuna(4) = lColunaI + 7.3
    lColuna(5) = lColunaI + 9.6
    lColuna(6) = lColunaI + 11.44
    lColuna(7) = lColunaI + 13.28
    lColuna(8) = lColunaI + 15.12
    lColuna(9) = lColunaI + 16.96
    lColuna(10) = lColunaI + 18.8
    lColuna(11) = lColunaI + 20.5
    
    'Coordenadas das Linhas da grade E (em centimetros)
    lLinha(0) = lLinhaI + 0
    lLinha(1) = lLinhaI + 0.31
    lLinha(2) = lLinhaI + 0.62
    lLinha(3) = lLinhaI + 0.93
    lLinha(4) = lLinhaI + 1.24
    lLinha(5) = lLinhaI + 1.55
    lLinha(6) = lLinhaI + 1.86
    lLinha(7) = lLinhaI + 2.17
    lLinha(8) = lLinhaI + 2.48
    lLinha(9) = lLinhaI + 2.79
    lLinha(10) = lLinhaI + 3.1
    lLinha(11) = lLinhaI + 3.41
    lLinha(12) = lLinhaI + 3.72
    lLinha(13) = lLinhaI + 4.03
    lLinha(14) = lLinhaI + 4.34
    lLinha(15) = lLinhaI + 4.65
    lLinha(16) = lLinhaI + 4.96
    lLinha(17) = lLinhaI + 5.24
    
    Printer.DrawWidth = 6
    Printer.Line (lColuna(0), lLinha(0))-(lColuna(3), lLinha(0)), RGB(lR, lG, lB)
    
    'Imprime os dados do cabeçalho da grade E
    Printer.FontBold = False
    Printer.FontSize = 6
    ImprimeCentralizado "Resumo Vendas", lColuna(0), lColuna(1), lLinha(0) + 0.06, lLocal
    ImprimeCentralizado "Quantidade", lColuna(1), lColuna(2), lLinha(0) + 0.21, lLocal
    ImprimeCentralizado "Valor", lColuna(2), lColuna(3), lLinha(0) + 0.21, lLocal
    ImprimeCentralizado "Caixa Final", lColuna(3), lColuna(7), lLinha(0) + 0.04, lLocal
    
    'Imprime 1ª linha horizontal da grade E
    Printer.Line (lColuna(3), lLinha(1))-(lColuna(7), lLinha(1)), RGB(lR, lG, lB)
    
    'Imprime 1ª linha de dados da grade E
    ImprimeCentralizado "Outros Produtos", lColuna(0), lColuna(1), lLinha(1) + 0.06, lLocal
    ImprimeCentralizado "Créditos", lColuna(3), lColuna(4), lLinha(1) + 0.04, lLocal
    ImprimeCentralizado "Débitos", lColuna(4), lColuna(5), lLinha(1) + 0.04, lLocal
    ImprimeCentralizado "Saldo Dia", lColuna(5), lColuna(6), lLinha(1) + 0.04, lLocal
    ImprimeCentralizado "Diferenças", lColuna(6), lColuna(7), lLinha(1) + 0.04, lLocal
    
    'Imprime 2ª linha horizontal da grade E
    Printer.Line (lColuna(0), lLinha(2))-(lColuna(7), lLinha(2)), RGB(lR, lG, lB)
    
    'Imprime 2ª linha de dados da grade E
    ImprimeTexto "Lavagem", lColuna(0), lColuna(1), lLinha(2) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(2) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(2) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(3), lColuna(4), lLinha(2) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(3), lColuna(4), lLinha(2) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(4), lColuna(5), lLinha(2) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(4), lColuna(5), lLinha(2) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(5), lColuna(6), lLinha(2) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(5), lColuna(6), lLinha(2) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(2) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(6), lColuna(7), lLinha(2) + 0.04, lLocal
    
    'Imprime 3ª linha horizontal da grade E
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(3))-(lColuna(3), lLinha(3)), RGB(lR, lG, lB)
    Printer.DrawWidth = 6
    Printer.Line (lColuna(3), lLinha(3))-(lColuna(7), lLinha(3)), RGB(lR, lG, lB)
    
    'Imprime 3ª linha de dados da grade E
    ImprimeTexto "Lubrificantes", lColuna(0), lColuna(1), lLinha(3) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(3) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(3) + 0.04, lLocal
    
    'Imprime 4ª linha horizontal da grade E
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(4))-(lColuna(3), lLinha(4)), RGB(lR, lG, lB)
    
    'Imprime 4ª linha de dados da grade E
    ImprimeTexto "Baldes", lColuna(0), lColuna(1), lLinha(4) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(4) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(4) + 0.04, lLocal
    
    'Imprime 5ª linha horizontal da grade E
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(5))-(lColuna(3), lLinha(5)), RGB(lR, lG, lB)
    
    'Imprime 5ª linha de dados da grade E
    ImprimeTexto "Filtros", lColuna(0), lColuna(1), lLinha(5) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(5) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(5) + 0.04, lLocal
    ImprimeCentralizado "Caixa Acumulado", lColuna(3), lColuna(7), lLinha(5) + 0.04, lLocal
    
    'Imprime 6ª linha horizontal da grade E
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(6))-(lColuna(3), lLinha(6)), RGB(lR, lG, lB)
    Printer.DrawWidth = 6
    Printer.Line (lColuna(3), lLinha(6))-(lColuna(7), lLinha(6)), RGB(lR, lG, lB)
    
    'Imprime 6ª linha de dados da grade E
    ImprimeTexto "Estopas", lColuna(0), lColuna(1), lLinha(6) + 0.04, lLocal
    ImprimeCentralizado "Créditos", lColuna(3), lColuna(4), lLinha(6) + 0.04, lLocal
    ImprimeCentralizado "Débitos", lColuna(4), lColuna(5), lLinha(6) + 0.04, lLocal
    ImprimeCentralizado "Saldo Dia", lColuna(5), lColuna(6), lLinha(6) + 0.04, lLocal
    ImprimeCentralizado "Diferenças", lColuna(6), lColuna(7), lLinha(6) + 0.04, lLocal
    
    'Imprime 7ª linha horizontal da grade E
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(7))-(lColuna(3), lLinha(7)), RGB(lR, lG, lB)
    Printer.DrawWidth = 6
    Printer.Line (lColuna(3), lLinha(7))-(lColuna(7), lLinha(7)), RGB(lR, lG, lB)
    
    'Imprime 7ª linha de dados da grade E
    ImprimeTexto "Acessórios", lColuna(0), lColuna(1), lLinha(7) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(7) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(7) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(3), lColuna(4), lLinha(7) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(3), lColuna(4), lLinha(7) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(4), lColuna(5), lLinha(7) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(4), lColuna(5), lLinha(7) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(5), lColuna(6), lLinha(7) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(5), lColuna(6), lLinha(7) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(7) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(6), lColuna(7), lLinha(7) + 0.04, lLocal
    
    'Imprime 8ª linha horizontal da grade E
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(8))-(lColuna(3), lLinha(8)), RGB(lR, lG, lB)
    Printer.DrawWidth = 6
    Printer.Line (lColuna(3), lLinha(8))-(lColuna(7), lLinha(8)), RGB(lR, lG, lB)
    
    'Imprime 8ª linha de dados da grade E
    ImprimeTexto "Gás Cozinha", lColuna(0), lColuna(1), lLinha(8) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(8) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(8) + 0.04, lLocal
    
    'Imprime 9ª linha horizontal da grade E
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(9))-(lColuna(3), lLinha(9)), RGB(lR, lG, lB)
    Printer.Line (lColuna(5), lLinha(9))-(lColuna(11), lLinha(9)), RGB(lR, lG, lB)
    
    'Imprime 9ª linha de dados da grade E
    ImprimeTexto "Outros", lColuna(0), lColuna(1), lLinha(9) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(9) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(9) + 0.04, lLocal
    ImprimeCentralizado "Custo", lColuna(5), lColuna(11), lLinha(9) + 0.04, lLocal
    
    'Imprime 10ª linha horizontal da grade E
    Printer.DrawWidth = 3
    Printer.Line (lColuna(0), lLinha(10))-(lColuna(3), lLinha(10)), RGB(lR, lG, lB)
    Printer.Line (lColuna(5), lLinha(10))-(lColuna(11), lLinha(10)), RGB(lR, lG, lB)
    
    'Imprime 10ª linha de dados da grade E
    ImprimeTexto "Aferição Gas/AH", lColuna(0), lColuna(1), lLinha(10) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(10) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(10) + 0.04, lLocal
    ImprimeTexto "Volume", lColuna(5), lColuna(6), lLinha(10) + 0.04, lLocal
    ImprimeCentralizado "Gasolina", lColuna(6), lColuna(8), lLinha(10) + 0.04, lLocal
    ImprimeCentralizado "Alcool", lColuna(8), lColuna(10), lLinha(10) + 0.04, lLocal
    ImprimeCentralizado "Totais", lColuna(10), lColuna(11), lLinha(10) + 0.04, lLocal
  
    'Imprime 11ª linha horizontal da grade E
    Printer.DrawWidth = 6
    Printer.Line (lColuna(0), lLinha(11))-(lColuna(3), lLinha(11)), RGB(lR, lG, lB)
    Printer.DrawWidth = 3
    Printer.Line (lColuna(4), lLinha(11))-(lColuna(11), lLinha(11)), RGB(lR, lG, lB)
    
    'Imprime 11ª linha de dados da grade E
    ImprimeTexto "Total", lColuna(0), lColuna(1), lLinha(11) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(11) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(11) + 0.04, lLocal
    ImprimeTexto "ESTOQUE INICIAL", lColuna(4), lColuna(5), lLinha(11) + 0.04, lLocal
    ImprimeValor "0,0", lColuna(5), lColuna(6), lLinha(11) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(11) + 0.04, lLocal
    ImprimeValor "0,0000", lColuna(6), lColuna(7), lLinha(11) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(7), lColuna(8), lLinha(11) + 0.04, lLocal
    ImprimeValor "0,00", lColuna(7), lColuna(8), lLinha(11) + 0.04, lLocal
    ImprimeValor "0,0", lColuna(8), lColuna(9), lLinha(11) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(9), lColuna(10), lLinha(11) + 0.04, lLocal
    ImprimeValor "0,0000", lColuna(9), lColuna(10), lLinha(11) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(10), lColuna(11), lLinha(11) + 0.04, lLocal
    ImprimeValor "0,00", lColuna(10), lColuna(11), lLinha(11) + 0.04, lLocal
    
    'Imprime 12ª linha horizontal da grade E
    Printer.DrawWidth = 6
    Printer.Line (lColuna(0), lLinha(12))-(lColuna(3), lLinha(12)), RGB(lR, lG, lB)
    Printer.DrawWidth = 3
    Printer.Line (lColuna(4), lLinha(12))-(lColuna(11), lLinha(12)), RGB(lR, lG, lB)
    
    'Imprime 12ª linha de dados da grade E
    ImprimeTexto "Acumulado / Lubrif / Acessorios / Gas", lColuna(0), lColuna(2), lLinha(12) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(2), lColuna(3), lLinha(12) + 0.04, lLocal
    ImprimeValor "000,0", lColuna(2), lColuna(3), lLinha(12) + 0.04, lLocal
    ImprimeTexto "COMPRAS", lColuna(4), lColuna(5), lLinha(12) + 0.04, lLocal
    ImprimeValor "0,0", lColuna(5), lColuna(6), lLinha(12) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(12) + 0.04, lLocal
    ImprimeValor "0,0000", lColuna(6), lColuna(7), lLinha(12) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(7), lColuna(8), lLinha(12) + 0.04, lLocal
    ImprimeValor "0,00", lColuna(7), lColuna(8), lLinha(12) + 0.04, lLocal
    ImprimeValor "0,0", lColuna(8), lColuna(9), lLinha(12) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(9), lColuna(10), lLinha(12) + 0.04, lLocal
    ImprimeValor "0,0000", lColuna(9), lColuna(10), lLinha(12) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(10), lColuna(11), lLinha(12) + 0.04, lLocal
    ImprimeValor "0,00", lColuna(10), lColuna(11), lLinha(12) + 0.04, lLocal
    
    'Imprime 13ª linha horizontal da grade E
    Printer.DrawWidth = 6
    Printer.Line (lColuna(0), lLinha(13))-(lColuna(3), lLinha(13)), RGB(lR, lG, lB)
    Printer.DrawWidth = 3
    Printer.Line (lColuna(4), lLinha(13))-(lColuna(11), lLinha(13)), RGB(lR, lG, lB)
    
    'Imprime 13ª linha de dados da grade E
    ImprimeTexto "ESTOQUE FINAL", lColuna(4), lColuna(5), lLinha(13) + 0.04, lLocal
    ImprimeValor "0,0", lColuna(5), lColuna(6), lLinha(13) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(13) + 0.04, lLocal
    ImprimeValor "0,0000", lColuna(6), lColuna(7), lLinha(13) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(7), lColuna(8), lLinha(13) + 0.04, lLocal
    ImprimeValor "0,00", lColuna(7), lColuna(8), lLinha(13) + 0.04, lLocal
    ImprimeValor "0,0", lColuna(8), lColuna(9), lLinha(13) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(9), lColuna(10), lLinha(13) + 0.04, lLocal
    ImprimeValor "0,0000", lColuna(9), lColuna(10), lLinha(13) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(10), lColuna(11), lLinha(13) + 0.04, lLocal
    ImprimeValor "0,00", lColuna(10), lColuna(11), lLinha(13) + 0.04, lLocal
    
    'Imprime 14ª linha horizontal da grade E
    Printer.DrawWidth = 6
    Printer.Line (lColuna(0), lLinha(14))-(lColuna(2), lLinha(14)), RGB(lR, lG, lB)
    Printer.DrawWidth = 3
    Printer.Line (lColuna(4), lLinha(14))-(lColuna(11), lLinha(14)), RGB(lR, lG, lB)
    
    'Imprime 14ª linha de dados da grade E
    ImprimeTexto "Acumulado Aferição", lColuna(0), lColuna(1), lLinha(14) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(1), lColuna(2), lLinha(14) + 0.04, lLocal
    ImprimeValor "(0,00)", lColuna(1), lColuna(2), lLinha(14) + 0.04, lLocal
    
    'Imprime 15ª linha horizontal da grade E
    Printer.DrawWidth = 6
    Printer.Line (lColuna(0), lLinha(15))-(lColuna(2), lLinha(15)), RGB(lR, lG, lB)
    Printer.DrawWidth = 3
    Printer.Line (lColuna(5), lLinha(15))-(lColuna(11), lLinha(15)), RGB(lR, lG, lB)
    
    'Imprime 15ª linha de dados da grade E
    ImprimeCentralizado "CMV Gasolina", lColuna(5), lColuna(6), lLinha(15) + 0.04, lLocal
    ImprimeCentralizado "CMV Álcool", lColuna(6), lColuna(7), lLinha(15) + 0.04, lLocal
    ImprimeCentralizado "CMV", lColuna(7), lColuna(8), lLinha(15) + 0.04, lLocal
    ImprimeCentralizado "LUCRO BRUTO", lColuna(8), lColuna(9), lLinha(15) + 0.04, lLocal
    ImprimeCentralizado "DESPESAS", lColuna(9), lColuna(10), lLinha(15) + 0.04, lLocal
    ImprimeCentralizado "LUCRO LÍQUIDO", lColuna(10), lColuna(11), lLinha(15) + 0.04, lLocal
    
    'Imprime 16ª linha horizontal da grade E
    Printer.DrawWidth = 3
    Printer.Line (lColuna(5), lLinha(16))-(lColuna(11), lLinha(16)), RGB(lR, lG, lB)
    
    'Imprime 16ª linha de dados da grade E
    ImprimeTexto "R$", lColuna(5), lColuna(6), lLinha(16) + 0.04, lLocal
    ImprimeValor "0,00", lColuna(5), lColuna(6), lLinha(16) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(6), lColuna(7), lLinha(16) + 0.04, lLocal
    ImprimeValor "0,00", lColuna(6), lColuna(7), lLinha(16) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(7), lColuna(8), lLinha(16) + 0.04, lLocal
    ImprimeValor "0,00", lColuna(7), lColuna(8), lLinha(16) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(8), lColuna(9), lLinha(16) + 0.04, lLocal
    ImprimeValor "0,00", lColuna(8), lColuna(9), lLinha(16) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(9), lColuna(10), lLinha(16) + 0.04, lLocal
    ImprimeValor "0,00", lColuna(9), lColuna(10), lLinha(16) + 0.04, lLocal
    ImprimeTexto "R$", lColuna(10), lColuna(11), lLinha(16) + 0.04, lLocal
    ImprimeValor "0,00", lColuna(10), lColuna(11), lLinha(16) + 0.04, lLocal
    
    'Imprime 17ª linha horizontal da grade E
    Printer.DrawWidth = 3
    Printer.Line (lColuna(5), lLinha(17))-(lColuna(11), lLinha(17)), RGB(lR, lG, lB)
    
    'Imprime linhas verticais da grade E
    Printer.DrawWidth = 6
    Printer.Line (lColuna(0), lLinha(0))-(lColuna(0), lLinha(13)), RGB(lR, lG, lB)
    Printer.Line (lColuna(1), lLinha(0))-(lColuna(1), lLinha(12)), RGB(lR, lG, lB)
    Printer.Line (lColuna(2), lLinha(0))-(lColuna(2), lLinha(13)), RGB(lR, lG, lB)
    Printer.Line (lColuna(3), lLinha(0))-(lColuna(3), lLinha(13)), RGB(lR, lG, lB)
    
    Printer.Line (lColuna(4), lLinha(1))-(lColuna(4), lLinha(3)), RGB(lR, lG, lB)
    Printer.Line (lColuna(5), lLinha(1))-(lColuna(5), lLinha(3)), RGB(lR, lG, lB)
    Printer.Line (lColuna(6), lLinha(1))-(lColuna(6), lLinha(3)), RGB(lR, lG, lB)
    Printer.Line (lColuna(7), lLinha(1))-(lColuna(7), lLinha(3)), RGB(lR, lG, lB)
    
    Printer.Line (lColuna(4), lLinha(6))-(lColuna(4), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(5), lLinha(6))-(lColuna(5), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(6), lLinha(6))-(lColuna(6), lLinha(8)), RGB(lR, lG, lB)
    Printer.Line (lColuna(7), lLinha(6))-(lColuna(7), lLinha(8)), RGB(lR, lG, lB)
    
    Printer.Line (lColuna(0), lLinha(14))-(lColuna(0), lLinha(15)), RGB(lR, lG, lB)
    Printer.Line (lColuna(1), lLinha(14))-(lColuna(1), lLinha(15)), RGB(lR, lG, lB)
    Printer.Line (lColuna(2), lLinha(14))-(lColuna(2), lLinha(15)), RGB(lR, lG, lB)
    Printer.DrawWidth = 3
    Printer.Line (lColuna(4), lLinha(11))-(lColuna(4), lLinha(14)), RGB(lR, lG, lB)
    Printer.Line (lColuna(5), lLinha(9))-(lColuna(5), lLinha(17)), RGB(lR, lG, lB)
    Printer.Line (lColuna(6), lLinha(10))-(lColuna(6), lLinha(17)), RGB(lR, lG, lB)
    Printer.Line (lColuna(7), lLinha(11))-(lColuna(7), lLinha(17)), RGB(lR, lG, lB)
    Printer.Line (lColuna(8), lLinha(10))-(lColuna(8), lLinha(17)), RGB(lR, lG, lB)
    Printer.Line (lColuna(9), lLinha(11))-(lColuna(9), lLinha(17)), RGB(lR, lG, lB)
    Printer.Line (lColuna(10), lLinha(10))-(lColuna(10), lLinha(17)), RGB(lR, lG, lB)
    Printer.Line (lColuna(11), lLinha(9))-(lColuna(11), lLinha(17)), RGB(lR, lG, lB)
    
    
End Sub
Private Sub PrintRoutine(objPrint As Object, Optional Proporcao As Double = 1)
    ' As dimensoes estao em polegadas
    
    
    
    'Coordenada inicial da linha da grade A (em centimetros)
    lLinhaI = (cp(0.3)) * Proporcao
    'Coordenada inicial da coluna da grade A (em centimetros)
    lColunaI = (cp(3.8)) * Proporcao
    
    'Coordenadas das colunas da grade A (em centimetros)
    lColuna(0) = lColunaI + 0 * Proporcao
    lColuna(1) = lColunaI + (cp(2.2)) * Proporcao
    lColuna(2) = lColunaI + (cp(3.9)) * Proporcao
    lColuna(3) = lColunaI + (cp(5.6)) * Proporcao
    lColuna(4) = lColunaI + (cp(7.3)) * Proporcao
    lColuna(5) = lColunaI + (cp(9.6)) * Proporcao
    lColuna(6) = lColunaI + (cp(11.44)) * Proporcao
    lColuna(7) = lColunaI + (cp(13.28)) * Proporcao
    lColuna(8) = lColunaI + (cp(15.12)) * Proporcao
    lColuna(9) = lColunaI + (cp(16.96)) * Proporcao
    lColuna(10) = lColunaI + (cp(18.8)) * Proporcao
    lColuna(11) = lColunaI + (cp(20.5)) * Proporcao
    
    'Coordenadas das Linhas da grade A (em centimetros)
    lLinha(0) = lLinhaI + 0 * Proporcao
    lLinha(1) = lLinhaI + (cp(1)) * Proporcao
    lLinha(2) = lLinhaI + (cp(1.62)) * Proporcao
    lLinha(3) = lLinhaI + (cp(1.93)) * Proporcao
    lLinha(4) = lLinhaI + (cp(2.24)) * Proporcao
    lLinha(5) = lLinhaI + (cp(2.55)) * Proporcao
    lLinha(6) = lLinhaI + (cp(2.86)) * Proporcao
    lLinha(7) = lLinhaI + (cp(3.17)) * Proporcao
    lLinha(8) = lLinhaI + (cp(3.48)) * Proporcao
    
    
    
    'Imprime os dados do cabeçalho do relatório
    objPrint.DrawWidth = 6 * Proporcao
    objPrint.ScaleMode = vbCentimeters
    objPrint.Line (lColuna(0), lLinha(0))-(lColuna(11), lLinha(0)), RGB(lR, lG, lB)
    objPrint.FontSize = 12 * Proporcao
    objPrint.FontBold = True
    ImprimeCentralizado "MOVIMENTAÇÃO DIÁRIA " & Trim(UCase(g_nome_empresa)) & " - " & UCase(Format(CDate(msk_data.Text), "mmmm")) & "/" & Format(CDate(msk_data.Text), "yyyy"), lColuna(0), lColuna(10), lLinha(0) + 0.3, lLocal
    '''ImprimeCentralizado Format(Format(CDate(msk_data.Text), "dd"), "00"), lColuna(10), lColuna(11), lLinha(0) + 0.3, lLocal
    objPrint.FontBold = False
    objPrint.Line (lColuna(0), lLinha(1))-(lColuna(11), lLinha(1)), RGB(lR, lG, lB)
    
    objPrint.Line (lColuna(0), lLinha(2))-(lColuna(11), lLinha(2)), RGB(lR, lG, lB)
    
    
    'Imprime linha horizontal da grade A
    objPrint.DrawWidth = 3 * Proporcao
    objPrint.Line (lColuna(0), lLinha(3))-(lColuna(4), lLinha(3)), RGB(lR, lG, lB)
    objPrint.Line (lColuna(10), lLinha(3))-(lColuna(11), lLinha(3)), RGB(lR, lG, lB)
    
    'Linhas verticais (grade A)
    objPrint.Line (lColuna(0), lLinha(0))-(lColuna(0), lLinha(8)), RGB(lR, lG, lB)
    objPrint.Line (lColuna(1), lLinha(1))-(lColuna(1), lLinha(8)), RGB(lR, lG, lB)
    objPrint.Line (lColuna(2), lLinha(1))-(lColuna(2), lLinha(8)), RGB(lR, lG, lB)
    objPrint.Line (lColuna(3), lLinha(1))-(lColuna(3), lLinha(8)), RGB(lR, lG, lB)
    objPrint.Line (lColuna(4), lLinha(1))-(lColuna(4), lLinha(8)), RGB(lR, lG, lB)
    objPrint.Line (lColuna(5), lLinha(1))-(lColuna(5), lLinha(8)), RGB(lR, lG, lB)
    objPrint.Line (lColuna(6), lLinha(1))-(lColuna(6), lLinha(8)), RGB(lR, lG, lB)
    objPrint.Line (lColuna(7), lLinha(1))-(lColuna(7), lLinha(8)), RGB(lR, lG, lB)
    objPrint.Line (lColuna(8), lLinha(1))-(lColuna(8), lLinha(8)), RGB(lR, lG, lB)
    objPrint.Line (lColuna(9), lLinha(1))-(lColuna(9), lLinha(8)), RGB(lR, lG, lB)
    objPrint.Line (lColuna(10), lLinha(0))-(lColuna(10), lLinha(8)), RGB(lR, lG, lB)
    objPrint.Line (lColuna(11), lLinha(0))-(lColuna(11), lLinha(8)), RGB(lR, lG, lB)
    
    
    
    Exit Sub
    
    
    
    ' Imprime alguns gráficos
    objPrint.Line (1, 1)-(1 + 6.5, 1 + 9), , B
    objPrint.Line (1.1, 2)-(1.1, 2)
    '''objPrint.PaintPicture picture2, 1.1, 1.1, 0.8, 0.8
    objPrint.Line (2.1, 1.2)-(2.1 + 5.2, 1.2 + 0.7), _
                   RGB(200, 200, 200), BF
    
    ' Imprime o titulo
    With objPrint
        .Font.name = "Arial"
        .CurrentX = 2.3
        .CurrentY = 1.3
        .FontSize = 35 * Proporcao
        objPrint.Print "Impressão no VB"
    End With
    
    ' Imprime alguns circulos
    Dim x As Single
    For x = 3 To 5.5 Step 0.2
        objPrint.Circle (x, 3.5), 0.75
    Next
    
    ' Imprime algum texto
    With objPrint
       .Font.name = "Courier New"
       .FontSize = 30 * Proporcao
       .CurrentX = 1.5
       .CurrentY = 5
       objPrint.Print "Visualizando Impressão"
    
       .FontSize = 24 * Proporcao
       .CurrentX = 1.5
       .CurrentY = 6.5
       objPrint.Print "Simples e fácil"
    
       .FontSize = 18 * Proporcao
       .CurrentX = 1.5
       .CurrentY = 8
       objPrint.Print "Agora basta imprimir ...."
    End With
End Sub
Private Sub Relatorio()
    'Monta a Select
    lSQl = ""
    lSQl = lSQl & "SELECT Data, Periodo, [Codigo da Bomba], Abertura, Encerrante, [Quantidade da Saida], "
    lSQl = lSQl & "       [Preco de Custo], [Preco de Venda], [Tipo de Combustivel]"
    lSQl = lSQl & "  FROM Movimento_Bomba"
    lSQl = lSQl & " WHERE Empresa = " & g_empresa
    lSQl = lSQl & "   AND Data = #" & Format(msk_data.Text, "mm/dd/yyyy") & "#"
    lSQl = lSQl & " ORDER BY Data, Periodo, [Codigo da Bomba]"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQl)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        ImpDados
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
    
    cmd_sair.SetFocus
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = RetiraGString(1)
    cmd_imprimir.SetFocus
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    If ValidaCampos Then
        If SelecionaImpressoraHP(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        End If
    End If
End Sub
Private Function ScalePicPreviewToPrinterInches(picPreview As PictureBox) As Double
    
    Dim Proporcao As Double ' Proporcao entre a impressora e a imagem
    Dim LRGap As Double, TBGap As Double
    Dim HeightProporcao As Double, WidthProporcao As Double
    Dim PgWidth As Double, PgHeight As Double
    Dim smtemp As Long
    
    ' Pega o tamanho fisico da pagina em polegadas ( Inches )
    PgWidth = Printer.Width / 1440
    PgHeight = Printer.Height / 1440
    
    ' Econtre o tamanho da area não imprimível
    smtemp = Printer.ScaleMode
    Printer.ScaleMode = vbInches
    'Printer.ScaleMode = vbCentimeters
    LRGap = (PgWidth - Printer.ScaleWidth) / 2
    TBGap = (PgHeight - Printer.ScaleHeight) / 2
    Printer.ScaleMode = smtemp
    
    picPreview.ScaleMode = vbInches
    'picPreview.ScaleMode = vbCentimeters
    
    HeightProporcao = picPreview.ScaleHeight / PgHeight
    WidthProporcao = picPreview.ScaleWidth / PgWidth
    
    If HeightProporcao < WidthProporcao Then
       Proporcao = HeightProporcao
       smtemp = picPreview.Container.ScaleMode
       picPreview.Container.ScaleMode = vbInches
       'picPreview.Container.ScaleMode = vbCentimeters
       picPreview.Width = PgWidth * Proporcao
       picPreview.Container.ScaleMode = smtemp
    Else
       Proporcao = WidthProporcao
       smtemp = picPreview.Container.ScaleMode
       picPreview.Container.ScaleMode = vbInches
       'picPreview.Container.ScaleMode = vbCentimeters
       picPreview.Height = PgHeight * Proporcao
       picPreview.Container.ScaleMode = smtemp
    End If
    
    ' Define as propriedades do picture box
    picPreview.Scale (0, 0)-(PgWidth, PgHeight)
    picPreview.Font.name = Printer.Font.name
    picPreview.FontSize = Printer.FontSize * Proporcao
    picPreview.ForeColor = Printer.ForeColor
    picPreview.Cls
    
    ScalePicPreviewToPrinterInches = Proporcao
End Function
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
    'ElseIf Not MovMapaResumo.LocalizarData(g_empresa, CDate(msk_data.Text)) Then
    '    MsgBox "Não há registro na data informada.", vbInformation, "Atenção!"
    '    msk_data.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    Dim xProporcao As Double
    xProporcao = ScalePicPreviewToPrinterInches(Picture1)
    PrintRoutine Picture1, xProporcao
    Call GravaAuditoria(1, Me.name, 6, "")
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data.SetFocus
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
        cmd_imprimir.SetFocus
    End If
End Sub
