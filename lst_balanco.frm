VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_balanco 
   Caption         =   "Emissão do Balanço"
   ClientHeight    =   2355
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   6675
   Icon            =   "lst_balanco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2355
   ScaleWidth      =   6675
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_balanco.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Visualiza o balanço."
      Top             =   1380
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2940
      Picture         =   "lst_balanco.frx":1A24
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprime o balanço."
      Top             =   1380
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4740
      Picture         =   "lst_balanco.frx":302E
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1380
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6435
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_balanco.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_balanco.frx":599A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5820
         Picture         =   "lst_balanco.frx":6C74
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4680
         TabIndex        =   8
         Top             =   660
         Width           =   1095
         _ExtentX        =   1931
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
         Width           =   1095
         _ExtentX        =   1931
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
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   3660
         TabIndex        =   7
         Top             =   720
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
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_balanco"
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
Dim l_venda_a(1 To 12) As Currency
Dim l_venda_aa(1 To 12) As Currency
Dim l_venda_d(1 To 12) As Currency
Dim l_venda_da(1 To 12) As Currency
Dim l_venda_g(1 To 12) As Currency
Dim l_venda_ga(1 To 12) As Currency
Dim l_custo_a(1 To 12) As Currency
Dim l_custo_aa(1 To 12) As Currency
Dim l_custo_d(1 To 12) As Currency
Dim l_custo_da(1 To 12) As Currency
Dim l_custo_g(1 To 12) As Currency
Dim l_custo_ga(1 To 12) As Currency
Dim l_custo_1(1 To 12) As Currency
Dim l_custo_2(1 To 12) As Currency
Dim l_custo_3(1 To 12) As Currency
Dim l_venda_1(1 To 12) As Currency
Dim l_venda_2(1 To 12) As Currency
Dim l_venda_3(1 To 12) As Currency
Dim l_despesa_fixa(1 To 12) As Currency
Dim l_despesa_variavel(1 To 12) As Currency
Dim l_cartao(1 To 12) As Currency
Dim l_PostoAki(1 To 12) As Currency
Dim l_investimento(1 To 12) As Currency
Dim l_cheque_devolvido(1 To 12) As Currency
Dim lValeFuncionario(1 To 12) As Currency
Dim lFaltaCaixa(1 To 12) As Currency
Dim lOutrasReceitas(1 To 12) As Currency
Dim l_nome_empresa As String
Dim lSQL As String
Dim lNomeEmpresaAtual As String

Dim rstBaixaPagar As adodb.Recordset
Dim rstEmpresa As adodb.Recordset
Dim rstMovimentoAfericao As adodb.Recordset
Dim rstMovimentoBomba As adodb.Recordset
Dim rstVendaLubrificante As adodb.Recordset
Dim rstPostAki As adodb.Recordset
Private MovCartaoCredito As New cMovimentoCartaoCredito
Private MovFaltaCaixa As New cMovimentoFaltaCaixa
Private Sub CalculaAfericao(ByVal pEmpresa As Integer)
    
    lSQL = ""
    lSQL = lSQL & " SELECT [Tipo de Combustivel], SUM([Preco de Venda] * Quantidade) AS TotalVenda, SUM([Preco de Custo] * Quantidade) AS TotalCusto"
    lSQL = lSQL & "   FROM Movimento_Afericao"
    lSQL = lSQL & "  WHERE Empresa = " & pEmpresa
    lSQL = lSQL & "    AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "    AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "  GROUP BY [Tipo de Combustivel]"
    lSQL = lSQL & "  ORDER BY [Tipo de Combustivel]"
    Set rstMovimentoAfericao = Conectar.RsConexao(lSQL)
    If rstMovimentoAfericao.RecordCount > 0 Then
        Do Until rstMovimentoAfericao.EOF
            If Trim(rstMovimentoAfericao![Tipo de Combustivel]) = "A" Then
                l_custo_a(pEmpresa) = l_custo_a(pEmpresa) - rstMovimentoAfericao!TotalCusto
                l_custo_a(12) = l_custo_a(12) - rstMovimentoAfericao!TotalCusto
                l_venda_a(pEmpresa) = l_venda_a(pEmpresa) - rstMovimentoAfericao!TotalVenda
                l_venda_a(12) = l_venda_a(12) - rstMovimentoAfericao!TotalVenda
            ElseIf Trim(rstMovimentoAfericao![Tipo de Combustivel]) = "AA" Then
                l_custo_aa(pEmpresa) = l_custo_aa(pEmpresa) - rstMovimentoAfericao!TotalCusto
                l_custo_aa(12) = l_custo_aa(12) - rstMovimentoAfericao!TotalCusto
                l_venda_aa(pEmpresa) = l_venda_aa(pEmpresa) - rstMovimentoAfericao!TotalVenda
                l_venda_aa(12) = l_venda_aa(12) - rstMovimentoAfericao!TotalVenda
            ElseIf Trim(rstMovimentoAfericao![Tipo de Combustivel]) = "D" Then
                l_custo_d(pEmpresa) = l_custo_d(pEmpresa) - rstMovimentoAfericao!TotalCusto
                l_custo_d(12) = l_custo_d(12) - rstMovimentoAfericao!TotalCusto
                l_venda_d(pEmpresa) = l_venda_d(pEmpresa) - rstMovimentoAfericao!TotalVenda
                l_venda_d(12) = l_venda_d(12) - rstMovimentoAfericao!TotalVenda
            ElseIf Trim(rstMovimentoAfericao![Tipo de Combustivel]) = "DA" Then
                l_custo_da(pEmpresa) = l_custo_da(pEmpresa) - rstMovimentoAfericao!TotalCusto
                l_custo_da(12) = l_custo_da(12) - rstMovimentoAfericao!TotalCusto
                l_venda_da(pEmpresa) = l_venda_da(pEmpresa) - rstMovimentoAfericao!TotalVenda
                l_venda_da(12) = l_venda_da(12) - rstMovimentoAfericao!TotalVenda
            ElseIf Trim(rstMovimentoAfericao![Tipo de Combustivel]) = "G" Then
                l_custo_g(pEmpresa) = l_custo_g(pEmpresa) - rstMovimentoAfericao!TotalCusto
                l_custo_g(12) = l_custo_g(12) - rstMovimentoAfericao!TotalCusto
                l_venda_g(pEmpresa) = l_venda_g(pEmpresa) - rstMovimentoAfericao!TotalVenda
                l_venda_g(12) = l_venda_g(12) - rstMovimentoAfericao!TotalVenda
            ElseIf Trim(rstMovimentoAfericao![Tipo de Combustivel]) = "GA" Then
                l_custo_ga(pEmpresa) = l_custo_ga(pEmpresa) - rstMovimentoAfericao!TotalCusto
                l_custo_ga(12) = l_custo_ga(12) - rstMovimentoAfericao!TotalCusto
                l_venda_ga(pEmpresa) = l_venda_ga(pEmpresa) - rstMovimentoAfericao!TotalVenda
                l_venda_ga(12) = l_venda_ga(12) - rstMovimentoAfericao!TotalVenda
            Else
                MsgBox "teste" & rstMovimentoAfericao![Tipo de Combustivel]
            End If
            rstMovimentoAfericao.MoveNext
        Loop
    End If
    rstMovimentoAfericao.Close
    Set rstMovimentoAfericao = Nothing
    
    
'    With tbl_movimento_afericao
'        .Seek ">=", i, CDate(msk_data_i), 0, 0, 0
'        If Not .NoMatch Then
'            Do Until .EOF
'                If !Empresa <> i Or !Data > CDate(msk_data_f) Then
'                    Exit Do
'                End If
'                Select Case Trim(![Tipo de Combustivel])
'                    Case "A"
'                        l_custo_a(i) = l_custo_a(i) - Format(!Quantidade * ![Preco de Custo], "#########0.00")
'                        l_custo_a(12) = l_custo_a(12) - Format(!Quantidade * ![Preco de Custo], "#########0.00")
'                        l_venda_a(i) = l_venda_a(i) - Format(!Quantidade * ![Preco de Venda], "#########0.00")
'                        l_venda_a(12) = l_venda_a(12) - Format(!Quantidade * ![Preco de Venda], "#########0.00")
'                    Case "AA"
'                        l_custo_aa(i) = l_custo_aa(i) - Format(!Quantidade * ![Preco de Custo], "#########0.00")
'                        l_custo_aa(12) = l_custo_aa(12) - Format(!Quantidade * ![Preco de Custo], "#########0.00")
'                        l_venda_aa(i) = l_venda_aa(i) - Format(!Quantidade * ![Preco de Venda], "#########0.00")
'                        l_venda_aa(12) = l_venda_aa(12) - Format(!Quantidade * ![Preco de Venda], "#########0.00")
'                    Case "D"
'                        l_custo_d(i) = l_custo_d(i) - Format(!Quantidade * ![Preco de Custo], "#########0.00")
'                        l_custo_d(12) = l_custo_d(12) - Format(!Quantidade * ![Preco de Custo], "#########0.00")
'                        l_venda_d(i) = l_venda_d(i) - Format(!Quantidade * ![Preco de Venda], "#########0.00")
'                        l_venda_d(12) = l_venda_d(12) - Format(!Quantidade * ![Preco de Venda], "#########0.00")
'                    Case "DA"
'                        l_custo_da(i) = l_custo_da(i) - Format(!Quantidade * ![Preco de Custo], "#########0.00")
'                        l_custo_da(12) = l_custo_da(12) - Format(!Quantidade * ![Preco de Custo], "#########0.00")
'                        l_venda_da(i) = l_venda_da(i) - Format(!Quantidade * ![Preco de Venda], "#########0.00")
'                        l_venda_da(12) = l_venda_da(12) - Format(!Quantidade * ![Preco de Venda], "#########0.00")
'                    Case "G"
'                        l_custo_g(i) = l_custo_g(i) - Format(!Quantidade * ![Preco de Custo], "#########0.00")
'                        l_custo_g(12) = l_custo_g(12) - Format(!Quantidade * ![Preco de Custo], "#########0.00")
'                        l_venda_g(i) = l_venda_g(i) - Format(!Quantidade * ![Preco de Venda], "#########0.00")
'                        l_venda_g(12) = l_venda_g(12) - Format(!Quantidade * ![Preco de Venda], "#########0.00")
'                    Case "GA"
'                        l_custo_ga(i) = l_custo_ga(i) - Format(!Quantidade * ![Preco de Custo], "#########0.00")
'                        l_custo_ga(12) = l_custo_ga(12) - Format(!Quantidade * ![Preco de Custo], "#########0.00")
'                        l_venda_ga(i) = l_venda_ga(i) - Format(!Quantidade * ![Preco de Venda], "#########0.00")
'                        l_venda_ga(12) = l_venda_ga(12) - Format(!Quantidade * ![Preco de Venda], "#########0.00")
'                End Select
'                .MoveNext
'            Loop
'        End If
'    End With
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set MovCartaoCredito = Nothing
    Set MovFaltaCaixa = Nothing
End Sub
Private Sub ZeraVariaveis()
Dim i As Integer
    lLinha = 0
    lPagina = 0
    l_nome_empresa = ""
    For i = 1 To 12
        l_venda_a(i) = 0
        l_venda_aa(i) = 0
        l_venda_d(i) = 0
        l_venda_da(i) = 0
        l_venda_g(i) = 0
        l_venda_ga(i) = 0
        l_custo_a(i) = 0
        l_custo_aa(i) = 0
        l_custo_d(i) = 0
        l_custo_da(i) = 0
        l_custo_g(i) = 0
        l_custo_ga(i) = 0
        l_custo_1(i) = 0
        l_venda_1(i) = 0
        l_custo_2(i) = 0
        l_venda_2(i) = 0
        l_custo_3(i) = 0
        l_venda_3(i) = 0
        l_despesa_fixa(i) = 0
        l_despesa_variavel(i) = 0
        l_investimento(i) = 0
        l_cartao(i) = 0
        l_PostoAki(i) = 0
        l_cheque_devolvido(i) = 0
        lValeFuncionario(i) = 0
        lFaltaCaixa(i) = 0
        lOutrasReceitas(i) = 0
    Next
End Sub
Private Sub Relatorio()
    Dim i As Integer
    ZeraVariaveis
    ImpCab
    Call LoopEmpresa
    Call ImpTotal
    BioImprime "@@Printer.EndDoc"
    BioFechaImprime
    g_string = lLocal & lNomeArquivo & "|@|Relatório do Balanço|@|"
    frm_preview.Show 1
    cmd_sair.SetFocus
End Sub
Private Sub ImpMovimentoBomba(ByVal pEmpresa As Integer)
   
    lSQL = ""
    lSQL = lSQL & " SELECT [Tipo de Combustivel], SUM([Preco de Venda] * [Quantidade da Saida]) AS TotalVenda, SUM([Preco de Custo] * [Quantidade da Saida]) AS TotalCusto"
    lSQL = lSQL & "   FROM Movimento_Bomba"
    lSQL = lSQL & "  WHERE Empresa = " & pEmpresa
    lSQL = lSQL & "    AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "    AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "  GROUP BY [Tipo de Combustivel]"
    lSQL = lSQL & "  ORDER BY [Tipo de Combustivel]"
    Set rstMovimentoBomba = Conectar.RsConexao(lSQL)
    If rstMovimentoBomba.RecordCount > 0 Then
        Do Until rstMovimentoBomba.EOF
            If Trim(rstMovimentoBomba![Tipo de Combustivel]) = "A" Then
                l_custo_a(pEmpresa) = l_custo_a(pEmpresa) + rstMovimentoBomba!TotalCusto
                l_custo_a(12) = l_custo_a(12) + rstMovimentoBomba!TotalCusto
                l_venda_a(pEmpresa) = l_venda_a(pEmpresa) + rstMovimentoBomba!TotalVenda
                l_venda_a(12) = l_venda_a(12) + rstMovimentoBomba!TotalVenda
            ElseIf Trim(rstMovimentoBomba![Tipo de Combustivel]) = "AA" Then
                l_custo_aa(pEmpresa) = l_custo_aa(pEmpresa) + rstMovimentoBomba!TotalCusto
                l_custo_aa(12) = l_custo_aa(12) + rstMovimentoBomba!TotalCusto
                l_venda_aa(pEmpresa) = l_venda_aa(pEmpresa) + rstMovimentoBomba!TotalVenda
                l_venda_aa(12) = l_venda_aa(12) + rstMovimentoBomba!TotalVenda
            ElseIf Trim(rstMovimentoBomba![Tipo de Combustivel]) = "D" Then
                l_custo_d(pEmpresa) = l_custo_d(pEmpresa) + rstMovimentoBomba!TotalCusto
                l_custo_d(12) = l_custo_d(12) + rstMovimentoBomba!TotalCusto
                l_venda_d(pEmpresa) = l_venda_d(pEmpresa) + rstMovimentoBomba!TotalVenda
                l_venda_d(12) = l_venda_d(12) + rstMovimentoBomba!TotalVenda
            ElseIf Trim(rstMovimentoBomba![Tipo de Combustivel]) = "DA" Then
                l_custo_da(pEmpresa) = l_custo_da(pEmpresa) + rstMovimentoBomba!TotalCusto
                l_custo_da(12) = l_custo_da(12) + rstMovimentoBomba!TotalCusto
                l_venda_da(pEmpresa) = l_venda_da(pEmpresa) + rstMovimentoBomba!TotalVenda
                l_venda_da(12) = l_venda_da(12) + rstMovimentoBomba!TotalVenda
            ElseIf Trim(rstMovimentoBomba![Tipo de Combustivel]) = "G" Then
                l_custo_g(pEmpresa) = l_custo_g(pEmpresa) + rstMovimentoBomba!TotalCusto
                l_custo_g(12) = l_custo_g(12) + rstMovimentoBomba!TotalCusto
                l_venda_g(pEmpresa) = l_venda_g(pEmpresa) + rstMovimentoBomba!TotalVenda
                l_venda_g(12) = l_venda_g(12) + rstMovimentoBomba!TotalVenda
            ElseIf Trim(rstMovimentoBomba![Tipo de Combustivel]) = "GA" Then
                l_custo_ga(pEmpresa) = l_custo_ga(pEmpresa) + rstMovimentoBomba!TotalCusto
                l_custo_ga(12) = l_custo_ga(12) + rstMovimentoBomba!TotalCusto
                l_venda_ga(pEmpresa) = l_venda_ga(pEmpresa) + rstMovimentoBomba!TotalVenda
                l_venda_ga(12) = l_venda_ga(12) + rstMovimentoBomba!TotalVenda
            Else
                MsgBox "teste" & rstMovimentoBomba![Tipo de Combustivel]
            End If
            rstMovimentoBomba.MoveNext
        Loop
    End If
    rstMovimentoBomba.Close
    Set rstMovimentoBomba = Nothing
    
'    With tbl_movimento_bomba
'        .Seek ">=", i, CDate(msk_data_i), "  ", 0, 0
'        If Not .NoMatch Then
'            Do Until .EOF
'                If !Empresa <> i Or !Data > CDate(msk_data_f) Then
'                    Exit Do
'                End If
'                x_custo = ![Preco de Custo] * ![Quantidade da Saida]
'                x_venda = ![Preco de Venda] * ![Quantidade da Saida]
'                If Trim(![Tipo de Combustivel]) = "A" Then
'                    l_custo_a(i) = l_custo_a(i) + x_custo
'                    l_custo_a(12) = l_custo_a(12) + x_custo
'                    l_venda_a(i) = l_venda_a(i) + x_venda
'                    l_venda_a(12) = l_venda_a(12) + x_venda
'                ElseIf Trim(![Tipo de Combustivel]) = "AA" Then
'                    l_custo_aa(i) = l_custo_aa(i) + x_custo
'                    l_custo_aa(12) = l_custo_aa(12) + x_custo
'                    l_venda_aa(i) = l_venda_aa(i) + x_venda
'                    l_venda_aa(12) = l_venda_aa(12) + x_venda
'                ElseIf Trim(![Tipo de Combustivel]) = "D" Then
'                    l_custo_d(i) = l_custo_d(i) + x_custo
'                    l_custo_d(12) = l_custo_d(12) + x_custo
'                    l_venda_d(i) = l_venda_d(i) + x_venda
'                    l_venda_d(12) = l_venda_d(12) + x_venda
'                ElseIf Trim(![Tipo de Combustivel]) = "DA" Then
'                    l_custo_da(i) = l_custo_da(i) + x_custo
'                    l_custo_da(12) = l_custo_da(12) + x_custo
'                    l_venda_da(i) = l_venda_da(i) + x_venda
'                    l_venda_da(12) = l_venda_da(12) + x_venda
'                ElseIf Trim(![Tipo de Combustivel]) = "G" Then
'                    l_custo_g(i) = l_custo_g(i) + x_custo
'                    l_custo_g(12) = l_custo_g(12) + x_custo
'                    l_venda_g(i) = l_venda_g(i) + x_venda
'                    l_venda_g(12) = l_venda_g(12) + x_venda
'                ElseIf Trim(![Tipo de Combustivel]) = "GA" Then
'                    l_custo_ga(i) = l_custo_ga(i) + x_custo
'                    l_custo_ga(12) = l_custo_ga(12) + x_custo
'                    l_venda_ga(i) = l_venda_ga(i) + x_venda
'                    l_venda_ga(12) = l_venda_ga(12) + x_venda
'                Else
'                    MsgBox "teste" & ![Tipo de Combustivel]
'                End If
'                .MoveNext
'            Loop
'        End If
'    End With
End Sub
Private Sub ImpMovimentoCartao(pEmpresa As Integer)
    Dim xValor As Currency
    xValor = MovCartaoCredito.TotalAdministrativoBaixado(pEmpresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text))
    l_cartao(pEmpresa) = l_cartao(pEmpresa) + xValor
    l_cartao(12) = l_cartao(12) + xValor
End Sub
Private Sub TotalizaValeFuncionario(pEmpresa As Integer)
    Dim xValor As Currency
    xValor = MovFaltaCaixa.TotalValeCaixa(pEmpresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), "0", 0)
    lValeFuncionario(pEmpresa) = lValeFuncionario(pEmpresa) + xValor
    lValeFuncionario(12) = lValeFuncionario(12) + xValor
End Sub
Private Sub TotalizaFaltaCaixa(pEmpresa As Integer)
    Dim xValor As Currency
    xValor = MovFaltaCaixa.TotalFaltaCaixa(pEmpresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), "0", 0)
    lFaltaCaixa(pEmpresa) = lFaltaCaixa(pEmpresa) + xValor
    lFaltaCaixa(12) = lFaltaCaixa(12) + xValor
End Sub
Private Sub ImpPostoAki(ByVal pEmpresa As Integer)

Const CODIGO_LANCAMENTO_PADRAO_POSTOAKI As Integer = 31

    lSQL = ""
    lSQL = lSQL & "SELECT sum([Valor]) AS Total"
    lSQL = lSQL & " FROM MovimentoCaixaPista"
    lSQL = lSQL & " WHERE Empresa = " & pEmpresa
    lSQL = lSQL & " AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & " AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & " and [Codigo do Lancamento Padrao] = " & CODIGO_LANCAMENTO_PADRAO_POSTOAKI
    
    Set rstPostAki = Conectar.RsConexao(lSQL)
    If rstPostAki.RecordCount > 0 Then
        If IsNull(rstPostAki!total) Then
            l_PostoAki(pEmpresa) = 0
        Else
            l_PostoAki(pEmpresa) = rstPostAki!total
            l_PostoAki(12) = l_PostoAki(12) + rstPostAki!total
        End If
    Else
        l_PostoAki(pEmpresa) = 0
    End If
    rstPostAki.Close
    Set rstPostAki = Nothing
End Sub
Private Sub ImpMovimentoLubrificante(ByVal pEmpresa As Integer)
    
    lSQL = ""
    lSQL = lSQL & " SELECT Grupo.Nome, SUM(Movimento_Lubrificante.[Valor Total]) AS TotalVenda, SUM(Movimento_Lubrificante.[Valor Custo] * Movimento_Lubrificante.Quantidade) AS TotalCusto, Produto.[Codigo do Grupo] AS CodigoGrupo"
    lSQL = lSQL & "   FROM Movimento_Lubrificante, Produto, Grupo"
    lSQL = lSQL & "  WHERE Movimento_Lubrificante.Empresa = " & pEmpresa
    lSQL = lSQL & "    AND Movimento_Lubrificante.Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "    AND Movimento_Lubrificante.Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "    AND Movimento_Lubrificante.[Codigo do Produto2] = Produto.Codigo"
    lSQL = lSQL & "    AND Produto.[Codigo do Grupo] = Grupo.Codigo"
    lSQL = lSQL & "  GROUP BY Produto.[Codigo do Grupo], Grupo.Nome"
    lSQL = lSQL & "  ORDER BY Produto.[Codigo do Grupo], Grupo.Nome"
    Set rstVendaLubrificante = Conectar.RsConexao(lSQL)
    If rstVendaLubrificante.RecordCount > 0 Then
        Do Until rstVendaLubrificante.EOF
            If rstVendaLubrificante!CodigoGrupo = 1 Then
                l_custo_1(pEmpresa) = l_custo_1(pEmpresa) + rstVendaLubrificante!TotalCusto
                l_custo_1(12) = l_custo_1(12) + rstVendaLubrificante!TotalCusto
                l_venda_1(pEmpresa) = l_venda_1(pEmpresa) + rstVendaLubrificante!TotalVenda
                l_venda_1(12) = l_venda_1(12) + rstVendaLubrificante!TotalVenda
            ElseIf rstVendaLubrificante!CodigoGrupo = 2 Then
                l_custo_2(pEmpresa) = l_custo_2(pEmpresa) + rstVendaLubrificante!TotalCusto
                l_custo_2(12) = l_custo_2(12) + rstVendaLubrificante!TotalCusto
                l_venda_2(pEmpresa) = l_venda_2(pEmpresa) + rstVendaLubrificante!TotalVenda
                l_venda_2(12) = l_venda_2(12) + rstVendaLubrificante!TotalVenda
            ElseIf rstVendaLubrificante!CodigoGrupo > 2 Then
                l_custo_3(pEmpresa) = l_custo_3(pEmpresa) + rstVendaLubrificante!TotalCusto
                l_custo_3(12) = l_custo_3(12) + rstVendaLubrificante!TotalCusto
                l_venda_3(pEmpresa) = l_venda_3(pEmpresa) + rstVendaLubrificante!TotalVenda
                l_venda_3(12) = l_venda_3(12) + rstVendaLubrificante!TotalVenda
            End If
            rstVendaLubrificante.MoveNext
        Loop
    End If
    rstVendaLubrificante.Close
    Set rstVendaLubrificante = Nothing
    
'    With tbl_movimento_lubrificante
'        .Seek ">=", i, CDate(msk_data_i.Text), 0, "0", 0, 0
'        If Not .NoMatch Then
'            Do Until .EOF
'                If !Empresa <> i Or !Data > CDate(msk_data_f) Then
'                    Exit Do
'                End If
'                tbl_produto.Seek "=", ![Codigo do Produto2]
'                If Not tbl_produto.NoMatch Then
'                    If tbl_produto![Codigo do Grupo] = 1 Then
'                        l_custo_1(i) = l_custo_1(i) + ![Valor Custo] * !Quantidade
'                        l_custo_1(12) = l_custo_1(12) + ![Valor Custo] * !Quantidade
'                        l_venda_1(i) = l_venda_1(i) + ![Valor Total]
'                        l_venda_1(12) = l_venda_1(12) + ![Valor Total]
'                    ElseIf tbl_produto![Codigo do Grupo] = 2 Then
'                        l_custo_2(i) = l_custo_2(i) + ![Valor Custo] * !Quantidade
'                        l_custo_2(12) = l_custo_2(12) + ![Valor Custo] * !Quantidade
'                        l_venda_2(i) = l_venda_2(i) + ![Valor Total]
'                        l_venda_2(12) = l_venda_2(12) + ![Valor Total]
'                    ElseIf tbl_produto![Codigo do Grupo] > 2 Then
'                        l_custo_3(i) = l_custo_3(i) + ![Valor Custo] * !Quantidade
'                        l_custo_3(12) = l_custo_3(12) + ![Valor Custo] * !Quantidade
'                        l_venda_3(i) = l_venda_3(i) + ![Valor Total]
'                        l_venda_3(12) = l_venda_3(12) + ![Valor Total]
'                    Else
'                        MsgBox "teste" & tbl_produto![Codigo do Grupo]
'                    End If
'                Else
'                    l_custo_3(i) = l_custo_3(i) + ![Valor Custo] * !Quantidade
'                    l_custo_3(12) = l_custo_3(12) + ![Valor Custo] * !Quantidade
'                    l_venda_3(i) = l_venda_3(i) + ![Valor Total]
'                    l_venda_3(12) = l_venda_3(12) + ![Valor Total]
'                End If
'                .MoveNext
'            Loop
'        End If
'    End With
End Sub
Private Sub LoopEmpresa()
    Dim i As Integer
    Dim x_venda As Currency
    Dim x_custo As Currency
    Dim x_venda2 As Currency
    Dim x_custo2 As Currency
    Dim x_despesa As Currency
    Dim x_lucro_liquido As Currency
    Dim x_linha As String
    
    lSQL = ""
    lSQL = lSQL & " SELECT Codigo, Nome"
    lSQL = lSQL & "   FROM Empresas"
    lSQL = lSQL & "  WHERE Codigo < " & 12
    lSQL = lSQL & "  ORDER BY Codigo"
    Set rstEmpresa = Conectar.RsConexao(lSQL)
    If rstEmpresa.RecordCount > 0 Then
        With rstEmpresa
            .MoveFirst
            Do Until .EOF
                lNomeEmpresaAtual = !Nome
                If !Codigo = 6 Or !Codigo = 11 Then
                    x_linha = "+--------------------------------+--------------------------------+--------------+--------------+--------------+--------+---------------+"
                    Mid(x_linha, 5, 22) = " Cerrado Informática. "
                    BioImprime "@Printer.Print " & x_linha
                    BioImprime "@@Printer.NewPage"
                    ImpCab
                End If
                If !Codigo > 11 Then
                    Exit Do
                End If
                i = !Codigo
                ImpMovimentoBomba i
                CalculaAfericao i
                x_venda = l_venda_a(i) + l_venda_aa(i) + l_venda_d(i) + l_venda_da(i) + l_venda_g(i) + l_venda_ga(i)
                x_custo = l_custo_a(i) + l_custo_aa(i) + l_custo_d(i) + l_custo_da(i) + l_custo_g(i) + l_custo_ga(i)
                Call ImpDet(Mid(!Nome, 1, 30), "Álcool Hidratado              ", l_venda_a(i), l_custo_a(i))
                Call ImpDet(Mid(!Nome, 1, 30), "Álcool Aditivado              ", l_venda_aa(i), l_custo_aa(i))
                Call ImpDet(Mid(!Nome, 1, 30), "Diesel                        ", l_venda_d(i), l_custo_d(i))
                Call ImpDet(Mid(!Nome, 1, 30), "Diesel Aditivado              ", l_venda_da(i), l_custo_da(i))
                Call ImpDet(Mid(!Nome, 1, 30), "Gasolina                      ", l_venda_g(i), l_custo_g(i))
                Call ImpDet(Mid(!Nome, 1, 30), "Gasolina Aditivada            ", l_venda_ga(i), l_custo_ga(i))
                Call ImpDet(Mid(!Nome, 1, 30), "** Sub-Total dos Combustíveis ", x_venda, x_custo)
                ImpMovimentoLubrificante i
                x_venda2 = l_venda_1(i) + l_venda_2(i) + l_venda_3(i)
                x_custo2 = l_custo_1(i) + l_custo_2(i) + l_custo_3(i)
                Call ImpDet(Mid(!Nome, 1, 30), "Óleos e Lubrificantes         ", l_venda_1(i), l_custo_1(i))
                Call ImpDet(Mid(!Nome, 1, 30), "Filtros                       ", l_venda_2(i), l_custo_2(i))
                Call ImpDet(Mid(!Nome, 1, 30), "Diversos                      ", l_venda_3(i), l_custo_3(i))
                Call ImpDet(Mid(!Nome, 1, 30), "** Sub-Total dos Produtos     ", x_venda2, x_custo2)
                x_venda2 = x_venda + x_venda2
                x_custo2 = x_custo + x_custo2
                Call ImpDet(Mid(!Nome, 1, 30), "** Total das Vendas           ", x_venda2, x_custo2)
                ImpBaixaPagar (i)
                ImpMovimentoCartao (i)
                ImpPostoAki (i)
                'TotalizaValeFuncionario (i)
                'TotalizaFaltaCaixa (i)
                x_despesa = l_despesa_fixa(i) + l_despesa_variavel(i) + l_cartao(i) + l_cheque_devolvido(i) + lValeFuncionario(i) + lFaltaCaixa(i) + l_PostoAki(i)
                Call ImpDetBaixaPagar(Mid(!Nome, 1, 30), "Despesas Fixas                ", l_despesa_fixa(i), "")
                Call ImpDetBaixaPagar(Mid(!Nome, 1, 30), "Despesas Variáveis            ", l_despesa_variavel(i), "")
                Call ImpDetBaixaPagar(Mid(!Nome, 1, 30), "Cheques Devolvidos            ", l_cheque_devolvido(i), "")
                Call ImpDetBaixaPagar(Mid(!Nome, 1, 30), "Administradoras de Cartões    ", l_cartao(i), "")
                Call ImpDetBaixaPagar(Mid(!Nome, 1, 30), "Desconto APP PostoAki         ", l_PostoAki(i), "")
                
                
                
                'Call ImpDetBaixaPagar(Mid(!Nome, 1, 30), "Vales de Funcionários         ", lValeFuncionario(i), "")
                'Call ImpDetBaixaPagar(Mid(!Nome, 1, 30), "Falta de Caixa                ", lFaltaCaixa(i), "")
                Call ImpDetBaixaPagar(Mid(!Nome, 1, 30), "** Total das Despesas         ", x_despesa, "")
                ImpOutrasReceitas (i)
                x_lucro_liquido = x_venda2 - x_custo2 - x_despesa + lOutrasReceitas(i)
                Call ImpDetBaixaPagar(Mid(!Nome, 1, 30), "*** Lucro Líquido             ", x_lucro_liquido, "")
                Call ImpDetBaixaPagar(Mid(!Nome, 1, 30), "Investimentos                 ", l_investimento(i), "")
                If l_investimento(i) > 0 Then
                    x_lucro_liquido = x_lucro_liquido - l_investimento(i)
                    Call ImpDetBaixaPagar(Mid(!Nome, 1, 30), "*** Lucro Líquido Capitalizado", x_lucro_liquido, "")
                End If
                If (x_venda2 > 0 Or x_despesa > 0) Then
                    BioImprime "@Printer.Print " & "+--------------------------------+--------------------------------+--------------+--------------+--------------+--------+---------------+"
                End If
                .MoveNext
            Loop
        End With
    End If
    rstEmpresa.Close
    Set rstEmpresa = Nothing
End Sub
Private Sub ImpTotal()
    Dim x_venda As Currency
    Dim x_venda2 As Currency
    Dim x_custo As Currency
    Dim x_custo2 As Currency
    Dim x_despesa As Currency
    Dim x_lucro_liquido As Currency
    Dim x_linha As String
    x_venda = l_venda_a(12) + l_venda_aa(12) + l_venda_d(12) + l_venda_da(12) + l_venda_g(12) + l_venda_ga(12)
    x_custo = l_custo_a(12) + l_custo_aa(12) + l_custo_d(12) + l_custo_da(12) + l_custo_g(12) + l_custo_ga(12)
    Call ImpDet("*** TODOS OS POSTOS ***", "Álcool Hidratado              ", l_venda_a(12), l_custo_a(12))
    Call ImpDet("*** TODOS OS POSTOS ***", "Álcool Aditivado              ", l_venda_aa(12), l_custo_aa(12))
    Call ImpDet("*** TODOS OS POSTOS ***", "Diesel                        ", l_venda_d(12), l_custo_d(12))
    Call ImpDet("*** TODOS OS POSTOS ***", "Diesel Aditivado              ", l_venda_da(12), l_custo_da(12))
    Call ImpDet("*** TODOS OS POSTOS ***", "Gasolina                      ", l_venda_g(12), l_custo_g(12))
    Call ImpDet("*** TODOS OS POSTOS ***", "Gasolina Aditivada            ", l_venda_ga(12), l_custo_ga(12))
    Call ImpDet("*** TODOS OS POSTOS ***", "** Total dos Combustíveis     ", x_venda, x_custo)
    x_venda2 = l_venda_1(12) + l_venda_2(12) + l_venda_3(12)
    x_custo2 = l_custo_1(12) + l_custo_2(12) + l_custo_3(12)
    Call ImpDet("*** TODOS OS POSTOS ***", "Óleos e Lubrificantes         ", l_venda_1(12), l_custo_1(12))
    Call ImpDet("*** TODOS OS POSTOS ***", "Filtros                       ", l_venda_2(12), l_custo_2(12))
    Call ImpDet("*** TODOS OS POSTOS ***", "Diversos                      ", l_venda_3(12), l_custo_3(12))
    Call ImpDet("*** TODOS OS POSTOS ***", "** Sub-Total dos Produtos     ", x_venda2, x_custo2)
    x_venda2 = x_venda + x_venda2
    x_custo2 = x_custo + x_custo2
    Call ImpDet("*** TODOS OS POSTOS ***", "** Total das Vendas           ", x_venda2, x_custo2)
    x_despesa = l_despesa_fixa(12) + l_despesa_variavel(12) + l_cartao(12) + l_cheque_devolvido(12) + lValeFuncionario(12) + lFaltaCaixa(12) + l_PostoAki(12)
    x_lucro_liquido = x_venda2 - x_custo2 - x_despesa + lOutrasReceitas(12)
    Call ImpDetBaixaPagar("*** TODOS OS POSTOS ***", "Despesas Fixas                ", l_despesa_fixa(12), "")
    Call ImpDetBaixaPagar("*** TODOS OS POSTOS ***", "Despesas Variáveis            ", l_despesa_variavel(12), "")
    Call ImpDetBaixaPagar("*** TODOS OS POSTOS ***", "Cheques Devolvidos            ", l_cheque_devolvido(12), "")
    Call ImpDetBaixaPagar("*** TODOS OS POSTOS ***", "Administradoras de Cartões    ", l_cartao(12), "")
    Call ImpDetBaixaPagar("*** TODOS OS POSTOS ***", "Desconto APP PostoAki         ", l_PostoAki(12), "")
    'Call ImpDetBaixaPagar("*** TODOS OS POSTOS ***", "Vales de Funcionários         ", lValeFuncionario(12),"")
    'Call ImpDetBaixaPagar("*** TODOS OS POSTOS ***", "Falta de Caixa                ", lFaltaCaixa(12),"")
    Call ImpDetBaixaPagar("*** TODOS OS POSTOS ***", "** Total das Despesas         ", x_despesa, "")
    Call ImpDetBaixaPagar("*** TODOS OS POSTOS ***", "** Outras Receitas            ", lOutrasReceitas(12), "")
    Call ImpDetBaixaPagar("*** TODOS OS POSTOS ***", "*** Lucro Líquido             ", x_lucro_liquido, "")
    Call ImpDetBaixaPagar("*** TODOS OS POSTOS ***", "Investimentos                 ", l_investimento(12), "")
    If l_investimento(12) > 0 Then
        x_lucro_liquido = x_lucro_liquido - l_investimento(12)
        Call ImpDetBaixaPagar("*** TODOS OS POSTOS ***", "*** Lucro Líquido Capitalizado", x_lucro_liquido, "")
    End If
    x_linha = "+--------------------------------+--------------------------------+--------------+--------------+--------------+--------+---------------+"
    Mid(x_linha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpDet(x_empresa As String, x_historico As String, x_venda As Currency, x_custo As Currency)
    Dim x_lucro As Currency
    Dim x_porc_1 As Currency
    Dim x_linha As String
    Dim i As Integer
    If x_venda > 0 Then
        If lLinha >= 60 Then
            x_linha = "+--------------------------------+--------------------------------+--------------+--------------+--------------+--------+---------------+"
            Mid(x_linha, 5, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & x_linha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        x_lucro = x_venda - x_custo
        x_porc_1 = x_lucro * 100 / x_venda
        '                   1         2         3         4         5         6         7         8         9        10        11        12        13     13
        '          12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
        x_linha = "|                                |                                |              |              |              |        |               |"
        If l_nome_empresa <> x_empresa Then
            l_nome_empresa = x_empresa
            Mid(x_linha, 3, 30) = x_empresa
        End If
        Mid(x_linha, 36, 30) = x_historico
        i = Len(Format(x_custo, "##,###,##0.00"))
        Mid(x_linha, 68 + 13 - i, i) = Format(x_custo, "##,###,##0.00")
        i = Len(Format(x_venda, "##,###,##0.00"))
        Mid(x_linha, 83 + 13 - i, i) = Format(x_venda, "##,###,##0.00")
        i = Len(Format(x_lucro, "##,###,##0.00"))
        Mid(x_linha, 98 + 13 - i, i) = Format(x_lucro, "##,###,##0.00")
        i = Len(Format(x_porc_1, "##0.00"))
        Mid(x_linha, 113 + 6 - i, i) = Format(x_porc_1, "##0.00")
        BioImprime "@@Printer.FontName = Sans Serif 17cpi"
        If Mid(x_historico, 1, 2) = "**" Then
            BioImprime "@@Printer.FontBold = True"
        End If
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.FontBold = False"
        lLinha = lLinha + 1
    End If
End Sub
Private Sub ImpDetBaixaPagar(ByVal pEmpresa As String, ByVal pHistorico As String, ByVal pValor As Currency, ByVal pObservacao As String)
    Dim x_linha As String
    Dim i As Integer
    
    If pValor <> 0 Then
        If lLinha >= 64 Then
            x_linha = "+--------------------------------+--------------------------------+--------------+--------------+--------------+--------+---------------+"
            Mid(x_linha, 5, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & x_linha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        x_linha = "|                                |                                |              |              |              |        |               |"
        If l_nome_empresa <> pEmpresa Then
            l_nome_empresa = pEmpresa
            Mid(x_linha, 3, 30) = pEmpresa
        End If
        Mid(x_linha, 36, 30) = pHistorico
        i = Len(Format(pValor, "##,###,##0.00;##,###,##0.00-"))
        If pValor > 0 Then
            Mid(x_linha, 98 + 13 - i, i) = Format(pValor, "##,###,##0.00")
        Else
            Mid(x_linha, 98 + 14 - i, i) = Format(pValor, "##,###,##0.00;##,###,##0.00-")
        End If
        Mid(x_linha, 122, 15) = pObservacao
        BioImprime "@@Printer.FontName = Sans Serif 17cpi"
        If Mid(pHistorico, 1, 2) = "**" Then
            BioImprime "@@Printer.FontBold = True"
        End If
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.FontBold = False"
        lLinha = lLinha + 1
    End If
End Sub
Private Sub ImpBaixaPagar(ByVal pEmpresa As Integer)
    
    lSQL = ""
    lSQL = lSQL & " SELECT codigo_conta, SUM(valor_pagamento) AS TotalPago"
    lSQL = lSQL & "   FROM Baixa_Pagar"
    lSQL = lSQL & "  WHERE Empresa = " & pEmpresa
    lSQL = lSQL & "    AND Data_Pagamento >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "    AND Data_Pagamento <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "    AND codigo_conta <= 5"
    lSQL = lSQL & "  GROUP BY codigo_conta"
    lSQL = lSQL & "  ORDER BY codigo_conta"
    Set rstBaixaPagar = Conectar.RsConexao(lSQL)
    If rstBaixaPagar.RecordCount > 0 Then
        Do Until rstBaixaPagar.EOF
            If rstBaixaPagar!codigo_conta = 2 Then
                l_despesa_fixa(pEmpresa) = l_despesa_fixa(pEmpresa) + rstBaixaPagar!TotalPago
                l_despesa_fixa(12) = l_despesa_fixa(12) + rstBaixaPagar!TotalPago
            ElseIf rstBaixaPagar!codigo_conta = 3 Then
                l_investimento(pEmpresa) = l_investimento(pEmpresa) + rstBaixaPagar!TotalPago
                l_investimento(12) = l_investimento(12) + rstBaixaPagar!TotalPago
            ElseIf rstBaixaPagar!codigo_conta = 4 Then
                l_despesa_variavel(pEmpresa) = l_despesa_variavel(pEmpresa) + rstBaixaPagar!TotalPago
                l_despesa_variavel(12) = l_despesa_variavel(12) + rstBaixaPagar!TotalPago
            ElseIf rstBaixaPagar!codigo_conta = 5 Then
                l_cheque_devolvido(pEmpresa) = l_cheque_devolvido(pEmpresa) + rstBaixaPagar!TotalPago
                l_cheque_devolvido(12) = l_cheque_devolvido(12) + rstBaixaPagar!TotalPago
            ElseIf rstBaixaPagar!codigo_conta > 5 Then
                MsgBox "Conta Não Configurada: " & rstBaixaPagar!codigo_conta
            End If
            rstBaixaPagar.MoveNext
        Loop
    End If
    rstBaixaPagar.Close
    Set rstBaixaPagar = Nothing
End Sub
Private Sub ImpOutrasReceitas(ByVal pEmpresa As Integer)
    
    lSQL = ""
    lSQL = lSQL & " SELECT Nome_Fornecedor, valor_pagamento, Complemento"
    lSQL = lSQL & "   FROM Baixa_Pagar"
    lSQL = lSQL & "  WHERE Empresa = " & pEmpresa
    lSQL = lSQL & "    AND Data_Pagamento >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "    AND Data_Pagamento <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "    AND codigo_conta = 6"
    lSQL = lSQL & "  ORDER BY Nome_Fornecedor"
    Set rstBaixaPagar = Conectar.RsConexao(lSQL)
    If rstBaixaPagar.RecordCount > 0 Then
        Do Until rstBaixaPagar.EOF
            Call ImpDetBaixaPagar(Mid(lNomeEmpresaAtual, 1, 30), rstBaixaPagar!Nome_Fornecedor, rstBaixaPagar!Valor_Pagamento, rstBaixaPagar!Complemento)
            lOutrasReceitas(pEmpresa) = lOutrasReceitas(pEmpresa) + rstBaixaPagar!Valor_Pagamento
            lOutrasReceitas(12) = lOutrasReceitas(12) + rstBaixaPagar!Valor_Pagamento
            rstBaixaPagar.MoveNext
        Loop
    End If
    rstBaixaPagar.Close
    Set rstBaixaPagar = Nothing
End Sub
Private Sub ImpCab()
    Dim x_linha As String
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
    BioImprime "@Printer.Print " & "+------------------------------------------------------------------------------+"
    BioImprime "@@Printer.FontBold = True"
    g_string = ReadINI("GRUPO DE EMPRESAS", "Nome do Grupo", gArquivoIni)
    x_linha = "| GRUPO X                                                          Página, " & Format(lPagina, "000") & " |"
    Mid(x_linha, 3, 40) = g_string
    g_string = ""
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| BALANÇO DOS POSTOS                                              , __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(x_linha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(x_linha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| REFERENTE AO PERIODO DE.: __/__/____ A __/__/____                            |"
    Mid(x_linha, 29, 10) = msk_data_i.Text
    Mid(x_linha, 42, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+--------------------------------+--------------------------------+--------------+--------------+--------------+--------+---------------+"
    BioImprime "@Printer.Print " & "| EMPRESA                        | CONTA                          | TOTAL  CUSTO | TOTAL  VENDA | LUCRO BRUTO  |%S/VENDA|COMPLEMENTO    |"
    BioImprime "@Printer.Print " & "+--------------------------------+--------------------------------+--------------+--------------+--------------+--------+---------------+"
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
        MsgBox "Data final deve ser maior ou igual a " & CDate(msk_data_i.Text) & ".", vbInformation, "Atenção!"
        msk_data_f.SetFocus
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
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        msk_data_i.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
        msk_data_f.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
        msk_data_i.SetFocus
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
Private Sub msk_data_f_GotFocus()
    msk_data_f.SelStart = 0
    msk_data_f.SelLength = 5
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
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
