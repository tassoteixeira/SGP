VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form emissao_analise_geral 
   Caption         =   "Emissão de Análise Geral"
   ClientHeight    =   2715
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   4890
   Icon            =   "lst_analise_geral.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2715
   ScaleWidth      =   4890
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3360
      Picture         =   "lst_analise_geral.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2040
      Picture         =   "lst_analise_geral.frx":199C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprime análise geral dos postos."
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   720
      Picture         =   "lst_analise_geral.frx":2FA6
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Visualiza análise geral dos postos."
      Top             =   1740
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1515
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4635
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   3480
         Picture         =   "lst_analise_geral.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   3480
         Picture         =   "lst_analise_geral.frx":599A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   3480
         Picture         =   "lst_analise_geral.frx":6C74
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   2340
         TabIndex        =   8
         Top             =   1080
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
         Left            =   2340
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
         Left            =   2340
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
         Left            =   780
         TabIndex        =   7
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   255
         Left            =   780
         TabIndex        =   4
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   255
         Left            =   780
         TabIndex        =   1
         Top             =   300
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_analise_geral"
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
Dim lChequePredatado(1 To 8) As Currency
Dim lChequeDevolvido(1 To 8) As Currency
Dim lChequeDevolvidoBaixado(1 To 8) As Currency
Dim lChequeDevolvidoBaixadoDescontado(1 To 8) As Currency
Dim lCartaoCredito(1 To 8) As Currency
Dim lCartaoCreditoHipercheque(1 To 8) As Currency
Dim lDuplicataReceberVencida(1 To 8) As Currency
Dim lDuplicataReceberVencer(1 To 8) As Currency
Dim lDuplicataReceberAfojac(1 To 8) As Currency
Dim lDuplicataReceberHipercheque(1 To 8) As Currency
Dim lNotaAbastecimento(1 To 8) As Currency
Dim lNotaAbastecimentoAfojac(1 To 8) As Currency
Dim lCombustivelA(1 To 8) As Currency
Dim lCombustivelAA(1 To 8) As Currency
Dim lCombustivelD(1 To 8) As Currency
Dim lCombustivelDA(1 To 8) As Currency
Dim lCombustivelG(1 To 8) As Currency
Dim lCombustivelGA(1 To 8) As Currency
Dim lEstoqueLubrificante(1 To 8) As Currency
Dim lEstoqueFiltro(1 To 8) As Currency
Dim lEstoqueDiverso(1 To 8) As Currency
Dim lPagar(1 To 8) As Currency
Dim lSubTotal(1 To 8) As Currency
Dim lTotal(1 To 8) As Currency
Dim lCodigoEmpresa(1 To 8) As Integer
Dim lNomeEmpresa(1 To 8) As String
Dim lSQL As String

Private BaixaChequeDevolvido As New cBaixaChequeDevolvido
Private BombaCombustivel As New cBomba
Private DuplicataReceber As New cDuplicataReceber
Private Estoque As New cEstoque
Private Grupo As New cGrupo
Private MedicaoCombustivel As New cMedicaoCombustivel
Private MovCartaoCredito As New cMovimentoCartaoCredito
Private MovCheque As New cMovimentoCheque
Private MovChequeDevolvido As New cMovimentoChequeDevolvido
Private MovContaPagar As New cMovimentoContaPagar
Private MovNotaAbastecimento As New cMovimentoNotaAbastecimento

Dim rstEmpresa As adodb.Recordset
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")

    Set BaixaChequeDevolvido = Nothing
    Set BombaCombustivel = Nothing
    Set DuplicataReceber = Nothing
    Set Estoque = Nothing
    Set Grupo = Nothing
    Set MedicaoCombustivel = Nothing
    Set MovCartaoCredito = Nothing
    Set MovCheque = Nothing
    Set MovChequeDevolvido = Nothing
    Set MovContaPagar = Nothing
    Set MovNotaAbastecimento = Nothing
    
    Set rstEmpresa = Nothing
End Sub
Private Sub ZeraVariaveis()
Dim i As Integer
    lLinha = 0
    lPagina = 0
    For i = 1 To 8
        lCodigoEmpresa(i) = 15
        lNomeEmpresa(i) = ""
        lChequePredatado(i) = 0
        lChequeDevolvido(i) = 0
        lChequeDevolvidoBaixado(i) = 0
        lChequeDevolvidoBaixadoDescontado(i) = 0
        lCartaoCredito(i) = 0
        lCartaoCreditoHipercheque(i) = 0
        lDuplicataReceberVencida(i) = 0
        lDuplicataReceberVencer(i) = 0
        lDuplicataReceberAfojac(i) = 0
        lDuplicataReceberHipercheque(i) = 0
        lNotaAbastecimento(i) = 0
        lNotaAbastecimentoAfojac(i) = 0
        lCombustivelA(i) = 0
        lCombustivelAA(i) = 0
        lCombustivelD(i) = 0
        lCombustivelDA(i) = 0
        lCombustivelG(i) = 0
        lCombustivelGA(i) = 0
        lEstoqueLubrificante(i) = 0
        lEstoqueFiltro(i) = 0
        lEstoqueDiverso(i) = 0
        lPagar(i) = 0
        lSubTotal(i) = 0
        lTotal(i) = 0
    Next
    For i = 1 To 8
        lCodigoEmpresa(i) = 15
        lNomeEmpresa(i) = ""
    Next
    g_string = ReadINI("GRUPO DE EMPRESAS", "Nome Abreviado das Empresas", gArquivoIni)
    
    
    lSQL = "SELECT Codigo, Nome, Inativo FROM Empresas ORDER BY Codigo"
    Set rstEmpresa = Conectar.RsConexao(lSQL)
    With rstEmpresa
        .MoveFirst
        i = 0
        Do Until .EOF
            If Not !Inativo Then
                i = i + 1
                If i > 8 Then
                    Exit Do
                End If
                lCodigoEmpresa(i) = !Codigo
                lNomeEmpresa(i) = RetiraGString(i)
            End If
            .MoveNext
        Loop
    End With
    g_string = ""
End Sub
Private Sub Relatorio()
    Dim i As Integer
    ZeraVariaveis
    ImpCab
    Call ImpChequePreDatado(True)
    Call ImpChequeDevolvido(True)
    'Call ImpChequeDevolvidoBaixado(True)
    'Call ImpChequeDevolvidoBaixadoDescontado(True)
    Call ImpCartaoCredito(True)
    For i = 1 To 8
        lSubTotal(i) = lSubTotal(i) + lChequePredatado(i)
        lSubTotal(i) = lSubTotal(i) + lCartaoCredito(i)
        'lSubTotal(i) = lSubTotal(i) + lCartaoCreditoHipercheque(i)
    Next
    Call ImpSubTotal(True)
    Call ImpDuplicataReceber(True)
    For i = 1 To 8
        lSubTotal(i) = lSubTotal(i) + lDuplicataReceberVencida(i)
        lSubTotal(i) = lSubTotal(i) + lDuplicataReceberVencer(i)
        lSubTotal(i) = lSubTotal(i) + lDuplicataReceberAfojac(i)
        lSubTotal(i) = lSubTotal(i) + lDuplicataReceberHipercheque(i)
    Next
    Call ImpSubTotal(True)
    Call ImpNotaAbastecimento(True)
    For i = 1 To 8
        lSubTotal(i) = lSubTotal(i) + lNotaAbastecimento(i)
        'lSubTotal(i) = lSubTotal(i) + lNotaAbastecimentoAfojac(i)
    Next
    Call ImpSubTotal(True)
    
    Call ImpEstoque(True)
    Call ImpSubTotal(True)
    
    Call ImpCombustivel(True)
    For i = 1 To 8
        lSubTotal(i) = lSubTotal(i) + lCombustivelA(i)
        lSubTotal(i) = lSubTotal(i) + lCombustivelAA(i)
        lSubTotal(i) = lSubTotal(i) + lCombustivelD(i)
        lSubTotal(i) = lSubTotal(i) + lCombustivelDA(i)
        lSubTotal(i) = lSubTotal(i) + lCombustivelG(i)
        lSubTotal(i) = lSubTotal(i) + lCombustivelGA(i)
    Next
    Call ImpSubTotal(True)
    For i = 1 To 8
        lTotal(i) = lTotal(i) + lChequePredatado(i)
        lTotal(i) = lTotal(i) + lCartaoCredito(i)
        lTotal(i) = lTotal(i) + lCartaoCreditoHipercheque(i)
        lTotal(i) = lTotal(i) + lDuplicataReceberVencida(i)
        lTotal(i) = lTotal(i) + lDuplicataReceberVencer(i)
        lTotal(i) = lTotal(i) + lDuplicataReceberAfojac(i)
        lTotal(i) = lTotal(i) + lDuplicataReceberHipercheque(i)
        lTotal(i) = lTotal(i) + lNotaAbastecimento(i)
        lTotal(i) = lTotal(i) + lNotaAbastecimentoAfojac(i)
        lTotal(i) = lTotal(i) + lCombustivelA(i)
        lTotal(i) = lTotal(i) + lCombustivelAA(i)
        lTotal(i) = lTotal(i) + lCombustivelD(i)
        lTotal(i) = lTotal(i) + lCombustivelDA(i)
        lTotal(i) = lTotal(i) + lCombustivelG(i)
        lTotal(i) = lTotal(i) + lCombustivelGA(i)
        lTotal(i) = lTotal(i) + lEstoqueLubrificante(i)
        lTotal(i) = lTotal(i) + lEstoqueFiltro(i)
        lTotal(i) = lTotal(i) + lEstoqueDiverso(i)
    Next
    Call ImpTotal(True)
    BioImprime "@Printer.Print " & "+----------------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+--------------+"
    Call ImpPagar(True)
    For i = 1 To 8
        lSubTotal(i) = lSubTotal(i) + lPagar(i)
    Next
    Call ImpSubTotal(True)
    
    For i = 1 To 8
        lTotal(i) = lTotal(i) - lPagar(i)
    Next
    Call ImpTotal(True)
    ImpRodape
    BioImprime "@@Printer.EndDoc"
    BioFechaImprime
    g_string = lLocal & lNomeArquivo & "|@|Relatório da Análise Geral dos Postos|@|"
    frm_preview.Show 1
    cmd_sair.SetFocus
End Sub
Private Sub ImpDuplicataReceber(ByVal pGeral As Boolean)
    Dim x_empresa As Integer

    For x_empresa = 1 To 7
        lDuplicataReceberVencida(x_empresa) = DuplicataReceber.TotalEntreDatas(x_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), CDate(Date) - 1, True)
        lDuplicataReceberVencida(8) = lDuplicataReceberVencida(8) + lDuplicataReceberVencida(x_empresa)
    Next

    For x_empresa = 1 To 7
        lDuplicataReceberVencer(x_empresa) = DuplicataReceber.TotalEntreDatas(x_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), CDate(Date) - 1, False)
        lDuplicataReceberVencer(8) = lDuplicataReceberVencer(8) + lDuplicataReceberVencer(x_empresa)
    Next

'    With tbl_duplicata_receber
'        If .RecordCount > 0 Then
'            .Index = "id_documento"
'            For x_empresa = 1 To 7
'                .Seek ">=", lCodigoEmpresa(x_empresa), 0
'                If Not .NoMatch Then
'                    Do Until .EOF
'                        If !Empresa <> lCodigoEmpresa(x_empresa) Then
'                            Exit Do
'                        End If
'                        If pGeral = True Or (pGeral = False And ![Data do Vencimento] >= CDate(msk_data_i) And ![Data do Vencimento] <= CDate(msk_data_f)) Then
'                            If ![Codigo do Cliente] = 26 Then
'                                lDuplicataReceberAfojac(x_empresa) = lDuplicataReceberAfojac(x_empresa) + ![Valor do Vencimento]
'                                lDuplicataReceberAfojac(8) = lDuplicataReceberAfojac(8) + ![Valor do Vencimento]
'                            ElseIf ![Codigo do Cliente] = 182 Then
'                                lDuplicataReceberHipercheque(x_empresa) = lDuplicataReceberHipercheque(x_empresa) + ![Valor do Vencimento]
'                                lDuplicataReceberHipercheque(8) = lDuplicataReceberHipercheque(8) + ![Valor do Vencimento]
'                            ElseIf ![Data do Vencimento] < CDate(msk_data) Then
'                                lDuplicataReceberVencida(x_empresa) = lDuplicataReceberVencida(x_empresa) + ![Valor do Vencimento]
'                                lDuplicataReceberVencida(8) = lDuplicataReceberVencida(8) + ![Valor do Vencimento]
'                            Else
'                                lDuplicataReceberVencer(x_empresa) = lDuplicataReceberVencer(x_empresa) + ![Valor do Vencimento]
'                                lDuplicataReceberVencer(8) = lDuplicataReceberVencer(8) + ![Valor do Vencimento]
'                            End If
'                        End If
'                        .MoveNext
'                    Loop
'                End If
'            Next
'        End If
'    End With
    Call ImpDet("MOV.TITULOS VENCIDOS", lDuplicataReceberVencida(1), lDuplicataReceberVencida(2), lDuplicataReceberVencida(3), lDuplicataReceberVencida(4), lDuplicataReceberVencida(5), lDuplicataReceberVencida(6), lDuplicataReceberVencida(7), lDuplicataReceberVencida(8))
    Call ImpDet("MOV.TITULOS A VENCER", lDuplicataReceberVencer(1), lDuplicataReceberVencer(2), lDuplicataReceberVencer(3), lDuplicataReceberVencer(4), lDuplicataReceberVencer(5), lDuplicataReceberVencer(6), lDuplicataReceberVencer(7), lDuplicataReceberVencer(8))
    'Call ImpDet("MOV.TITULOS AFOJAC  ", lDuplicataReceberAfojac(1), lDuplicataReceberAfojac(2), lDuplicataReceberAfojac(3), lDuplicataReceberAfojac(4), lDuplicataReceberAfojac(5), lDuplicataReceberAfojac(6), lDuplicataReceberAfojac(7), lDuplicataReceberAfojac(8))
    'Call ImpDet("MOV.TITULOS HIPERCHE", lDuplicataReceberHipercheque(1), lDuplicataReceberHipercheque(2), lDuplicataReceberHipercheque(3), lDuplicataReceberHipercheque(4), lDuplicataReceberHipercheque(5), lDuplicataReceberHipercheque(6), lDuplicataReceberHipercheque(7), lDuplicataReceberHipercheque(8))
End Sub
Private Sub ImpEstoque(ByVal pGeral As Boolean)
    Dim x_empresa As Integer
    Dim rstEstoque As adodb.Recordset
    Dim rstGrupo As adodb.Recordset
    Dim xValor As Currency
    Dim xEmpresa As Integer
    Dim i As Integer
    
    lSQL = "SELECT Codigo, Nome"
    lSQL = lSQL & "  FROM Grupo"
    lSQL = lSQL & " WHERE Codigo <> 4"
    lSQL = lSQL & " ORDER BY Nome"
    Set rstGrupo = Conectar.RsConexao(lSQL)
    If rstGrupo.RecordCount > 0 Then
        rstGrupo.MoveFirst
        For xEmpresa = 1 To 7
            Do Until rstGrupo.EOF
                lSQL = "SELECT Produto.[Codigo do Grupo], Sum(Estoque.Quantidade * Produto.[Preco de Custo]) As Total"
                lSQL = lSQL & "  FROM Estoque"
                lSQL = lSQL & "  LEFT JOIN Produto On Estoque.[Codigo do Produto2] = Produto.Codigo"
                lSQL = lSQL & "  LEFT JOIN Grupo On Estoque.[Grupo do Produto] = Grupo.Codigo"
                lSQL = lSQL & " WHERE Estoque.Empresa = " & lCodigoEmpresa(xEmpresa)
                lSQL = lSQL & "   AND Produto.[Codigo do Grupo] = " & rstGrupo("Codigo").Value
                lSQL = lSQL & "   AND Estoque.Quantidade > 0"
                lSQL = lSQL & " GROUP BY Produto.[Codigo do Grupo]"
                lSQL = lSQL & " ORDER BY Produto.[Codigo do Grupo]"
                Set rstEstoque = Conectar.RsConexao(lSQL)
                If rstEstoque.RecordCount > 0 Then
                    rstEstoque.MoveFirst
                    Do Until rstEstoque.EOF
                        If Not IsNull(rstEstoque("Total").Value) Then
                            lEstoqueDiverso(xEmpresa) = rstEstoque("Total").Value
                            lEstoqueDiverso(8) = lEstoqueDiverso(8) + lEstoqueDiverso(xEmpresa)
                            lSubTotal(xEmpresa) = lSubTotal(xEmpresa) + lEstoqueDiverso(xEmpresa)
                            lSubTotal(8) = lSubTotal(8) + lEstoqueDiverso(xEmpresa)
                        End If
                        rstEstoque.MoveNext
                    Loop
                End If
                Set rstEstoque = Nothing
                If lEstoqueDiverso(8) > 0 Then
                    Call ImpDet(rstGrupo("Nome").Value, lEstoqueDiverso(1), lEstoqueDiverso(2), lEstoqueDiverso(3), lEstoqueDiverso(4), lEstoqueDiverso(5), lEstoqueDiverso(6), lEstoqueDiverso(7), lEstoqueDiverso(8))
                End If
                For i = 1 To 8
                    lEstoqueDiverso(i) = 0
                Next
                rstGrupo.MoveNext
            Loop
        Next
    End If
    Set rstGrupo = Nothing
End Sub
Private Sub ImpNotaAbastecimento(ByVal pGeral As Boolean)
    Dim x_empresa As Integer
    
    For x_empresa = 1 To 7
        lNotaAbastecimento(x_empresa) = MovNotaAbastecimento.TotalData(x_empresa, 0, CDate(msk_data_i.Text), CDate(msk_data_f.Text), False)
        lNotaAbastecimento(8) = lNotaAbastecimento(8) + lNotaAbastecimento(x_empresa)
    Next
    
'    With tbl_movimento_nota_abastecimento
'        If .RecordCount > 0 Then
'            .Index = "id_cliente"
'            For x_empresa = 1 To 7
'                .Seek ">=", lCodigoEmpresa(x_empresa), 0, CDate("01/01/1900"), 0, 0, " "
'                If Not .NoMatch Then
'                    Do Until .EOF
'                        If !Empresa <> lCodigoEmpresa(x_empresa) Then
'                            Exit Do
'                        End If
'                        If pGeral = True Or (pGeral = False And ![Data do Abastecimento] >= CDate(msk_data_i) And ![Data do Abastecimento] <= CDate(msk_data_f)) Then
'                            If ![Codigo do Cliente] = 26 Then
'                                lNotaAbastecimentoAfojac(x_empresa) = lNotaAbastecimentoAfojac(x_empresa) + ![Valor Total]
'                                lNotaAbastecimentoAfojac(8) = lNotaAbastecimentoAfojac(8) + ![Valor Total]
'                            Else
'                                lNotaAbastecimento(x_empresa) = lNotaAbastecimento(x_empresa) + ![Valor Total]
'                                lNotaAbastecimento(8) = lNotaAbastecimento(8) + ![Valor Total]
'                            End If
'                        End If
'                        .MoveNext
'                    Loop
'                End If
'            Next
'        End If
'    End With
    Call ImpDet("NOTAS               ", lNotaAbastecimento(1), lNotaAbastecimento(2), lNotaAbastecimento(3), lNotaAbastecimento(4), lNotaAbastecimento(5), lNotaAbastecimento(6), lNotaAbastecimento(7), lNotaAbastecimento(8))
'    Call ImpDet("NOTAS AFOJAC        ", lNotaAbastecimentoAfojac(1), lNotaAbastecimentoAfojac(2), lNotaAbastecimentoAfojac(3), lNotaAbastecimentoAfojac(4), lNotaAbastecimentoAfojac(5), lNotaAbastecimentoAfojac(6), lNotaAbastecimentoAfojac(7), lNotaAbastecimentoAfojac(8))
End Sub
Function CalculaHipercheque(x_empresa As Integer)
'    Dim x_data_inicial As Date
'    x_data_inicial = "01/01/1900"
'    With tbl_duplicata_receber
'        .Index = "id_cliente_vencimento"
'        If .RecordCount > 0 Then
'            .Seek ">", 182, x_data_inicial, 0
'            If Not .NoMatch Then
'                Do Until .EOF
'                    If ![Codigo do Cliente] <> 182 Then
'                        Exit Do
'                    End If
'                    If !Empresa = x_empresa Then
'                        x_data_inicial = ![Data do Periodo Final] + 1
'                    End If
'                    .MoveNext
'                Loop
'            End If
'        End If
'    End With
'    With tbl_movimento_cartao_credito
'        .Index = "id_data_emissao"
'        If .RecordCount > 0 Then
'            .Seek ">", x_empresa, CDate(x_data_inicial), 0, 0
'            If Not .NoMatch Then
'                Do Until .EOF
'                    If !Empresa <> x_empresa Then
'                        Exit Do
'                    End If
'                    If !Periodo >= cbo_periodo_i And !Periodo <= cbo_periodo_f Then
'                        If ![Codigo do Cartao] = 4 Then
'                            l_hipercheque(x_empresa) = l_hipercheque(x_empresa) + !Valor
'                            l_hipercheque(12) = l_hipercheque(12) + !Valor
'                        End If
'                    End If
'                    .MoveNext
'                Loop
'            End If
'        End If
'    End With
End Function
Private Sub ImpCartaoCredito(ByVal pGeral As Boolean)
    Dim x_empresa As Integer
    'Dim x_data(1 To 7) As Date
    
    For x_empresa = 1 To 7
        lCartaoCredito(x_empresa) = MovCartaoCredito.TotalEntreDatas(x_empresa, False, True, CDate(msk_data_i.Text), CDate(msk_data_f.Text), 0)
        lCartaoCredito(8) = lCartaoCredito(8) + lCartaoCredito(x_empresa)
    Next
    
'    With tbl_duplicata_receber
'        .Index = "id_cliente_vencimento"
'        For x_empresa = 1 To 7
'            x_data(x_empresa) = CDate("01/01/1900")
'            If .RecordCount > 0 Then
'                .Seek "<=", 182, CDate("31/12/2500"), 9
'                If Not .NoMatch Then
'                    Do Until .BOF
'                        If ![Codigo do Cliente] <> 182 Then
'                            Exit Do
'                        End If
'                        If !Empresa = lCodigoEmpresa(x_empresa) Then
'                            x_data(x_empresa) = ![Data do Periodo Final] + 1
'                            Exit Do
'                        End If
'                        .MovePrevious
'                    Loop
'                End If
'            End If
'        Next
'    End With
'    'Teste porque nao esta mais vendendo com cartao
'    For x_empresa = 1 To 7
'        x_data(x_empresa) = CDate("31/12/2500")
'    Next
'    'Calcula cartão Hipercheque
'    With tbl_movimento_cartao_credito
'        If .RecordCount > 0 Then
'            .Index = "id_data_emissao"
'            For x_empresa = 1 To 7
'                .Seek ">=", lCodigoEmpresa(x_empresa), x_data(x_empresa), 0, 0
'                If Not .NoMatch Then
'                    Do Until .EOF
'                        If !Empresa <> lCodigoEmpresa(x_empresa) Then
'                            Exit Do
'                        End If
'                        If ![Codigo do Cartao] = 4 Then
'                            lCartaoCreditoHipercheque(x_empresa) = lCartaoCreditoHipercheque(x_empresa) + !Valor
'                            lCartaoCreditoHipercheque(8) = lCartaoCreditoHipercheque(8) + !Valor
'                        End If
'                        .MoveNext
'                    Loop
'                End If
'            Next
'        End If
'    End With
'    With tbl_movimento_cartao_credito
'        If .RecordCount > 0 Then
'            .Index = "id_data_vencimento2"
'            For x_empresa = 1 To 7
'                .Seek ">=", lCodigoEmpresa(x_empresa), CDate(msk_data), 0, 0, CDate("01/01/1900")
'                If Not .NoMatch Then
'                    Do Until .EOF
'                        If !Empresa <> lCodigoEmpresa(x_empresa) Then
'                            Exit Do
'                        End If
'                        If ![Codigo do Cartao] <> 4 Then
'                            lCartaoCredito(x_empresa) = lCartaoCredito(x_empresa) + !Valor
'                            lCartaoCredito(8) = lCartaoCredito(8) + !Valor
'                        End If
'                        .MoveNext
'                    Loop
'                End If
'            Next
'        End If
'    End With
    Call ImpDet("CARTÕES DE CRÉDITO  ", lCartaoCredito(1), lCartaoCredito(2), lCartaoCredito(3), lCartaoCredito(4), lCartaoCredito(5), lCartaoCredito(6), lCartaoCredito(7), lCartaoCredito(8))
    'Call ImpDet("CARTÕES HIPERCHEQUE ", lCartaoCreditoHipercheque(1), lCartaoCreditoHipercheque(2), lCartaoCreditoHipercheque(3), lCartaoCreditoHipercheque(4), lCartaoCreditoHipercheque(5), lCartaoCreditoHipercheque(6), lCartaoCreditoHipercheque(7), lCartaoCreditoHipercheque(8))
End Sub
Private Sub ImpChequeDevolvido(ByVal pGeral As Boolean)
    Dim x_empresa As Integer
    
    For x_empresa = 1 To 7
        lChequeDevolvido(x_empresa) = MovChequeDevolvido.TotalDataDevolucao(x_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text))
        lChequeDevolvido(8) = lChequeDevolvido(8) + lChequeDevolvido(x_empresa)
    Next
'    With tbl_movimento_cheque_devolvido
'        If .RecordCount > 0 Then
'            .Index = "id_data_emissao"
'            '+Empresa;+Data de Emissao;+Numero da Conta;+Numero do Cheque
'            For x_empresa = 1 To 7
'                .Seek ">=", lCodigoEmpresa(x_empresa), CDate("01/01/1900"), "          ", "      "
'                If Not .NoMatch Then
'                    Do Until .EOF
'                        If !Empresa <> lCodigoEmpresa(x_empresa) Then
'                            Exit Do
'                        End If
'                        If pGeral = True Or (pGeral = False And ![Data da Devolucao] >= CDate(msk_data_i) And ![Data da Devolucao] <= CDate(msk_data_f)) Then
'                            lChequeDevolvido(x_empresa) = lChequeDevolvido(x_empresa) + !Valor
'                            lChequeDevolvido(8) = lChequeDevolvido(8) + !Valor
'                        End If
'                        .MoveNext
'                    Loop
'                End If
'            Next
'        End If
'    End With
    Call ImpDet("CH.DEVOLVIDO        ", lChequeDevolvido(1), lChequeDevolvido(2), lChequeDevolvido(3), lChequeDevolvido(4), lChequeDevolvido(5), lChequeDevolvido(6), lChequeDevolvido(7), lChequeDevolvido(8))
End Sub
Private Sub ImpChequeDevolvidoBaixado(ByVal pGeral As Boolean)
    Dim x_empresa As Integer
    
    For x_empresa = 1 To 7
        lChequeDevolvidoBaixado(x_empresa) = BaixaChequeDevolvido.TotalDataPagamento(x_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text))
        lChequeDevolvidoBaixado(8) = lChequeDevolvidoBaixado(8) + lChequeDevolvidoBaixado(x_empresa)
    Next
'    With tbl_baixa_cheque_devolvido
'        If .RecordCount > 0 Then
'            .Index = "id_data_pagamento"
'            For x_empresa = 1 To 7
'                .Seek ">=", lCodigoEmpresa(x_empresa), CDate("01/01/1900"), 0, "    ", "      "
'                If Not .NoMatch Then
'                    Do Until .EOF
'                        If !Empresa <> lCodigoEmpresa(x_empresa) Then
'                            Exit Do
'                        End If
'                        If pGeral = True Or (pGeral = False And ![Data do Pagamento] >= CDate(msk_data_i) And ![Data do Pagamento] <= CDate(msk_data_f)) Then
'                            lChequeDevolvidoBaixado(x_empresa) = lChequeDevolvidoBaixado(x_empresa) + !Valor
'                            lChequeDevolvidoBaixado(8) = lChequeDevolvidoBaixado(8) + !Valor
'                        End If
'                        .MoveNext
'                    Loop
'                End If
'            Next
'        End If
'    End With
    Call ImpDet("CH.DEVOLVIDO BAIXADO", lChequeDevolvidoBaixado(1), lChequeDevolvidoBaixado(2), lChequeDevolvidoBaixado(3), lChequeDevolvidoBaixado(4), lChequeDevolvidoBaixado(5), lChequeDevolvidoBaixado(6), lChequeDevolvidoBaixado(7), lChequeDevolvidoBaixado(8))
End Sub
'Private Sub ImpChequeDevolvidoBaixadoDescontado(ByVal pGeral As Boolean)
'    Dim x_empresa As Integer
'    With tbl_baixa_cheque_devolvido_descontado
'        If .RecordCount > 0 Then
'            .Index = "id_data_pagamento"
'            For x_empresa = 1 To 7
'                .Seek ">=", lCodigoEmpresa(x_empresa), CDate("01/01/1900"), 0, "    ", "      "
'                If Not .NoMatch Then
'                    Do Until .EOF
'                        If !Empresa <> lCodigoEmpresa(x_empresa) Then
'                            Exit Do
'                        End If
'                        If pGeral = True Or (pGeral = False And ![Data do Pagamento] >= CDate(msk_data_i) And ![Data do Pagamento] <= CDate(msk_data_f)) Then
'                            lChequeDevolvidoBaixadoDescontado(x_empresa) = lChequeDevolvidoBaixadoDescontado(x_empresa) + !Valor
'                            lChequeDevolvidoBaixadoDescontado(8) = lChequeDevolvidoBaixadoDescontado(8) + !Valor
'                        End If
'                        .MoveNext
'                    Loop
'                End If
'            Next
'        End If
'    End With
'    Call ImpDet("CH.DEVOL.BAIX.DESC. ", lChequeDevolvidoBaixadoDescontado(1), lChequeDevolvidoBaixadoDescontado(2), lChequeDevolvidoBaixadoDescontado(3), lChequeDevolvidoBaixadoDescontado(4), lChequeDevolvidoBaixadoDescontado(5), lChequeDevolvidoBaixadoDescontado(6), lChequeDevolvidoBaixadoDescontado(7), lChequeDevolvidoBaixadoDescontado(8))
'End Sub
Private Sub ImpChequePreDatado(ByVal pGeral As Boolean)
    Dim x_empresa As Integer
    
    For x_empresa = 1 To 7
        lChequePredatado(x_empresa) = MovCheque.TotalEmissaoPeriodo(x_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), "1", "9", "0", "P")
        lChequePredatado(8) = lChequePredatado(8) + lChequePredatado(x_empresa)
    Next
    Call ImpDet("CH.PRÉ-DATADO       ", lChequePredatado(1), lChequePredatado(2), lChequePredatado(3), lChequePredatado(4), lChequePredatado(5), lChequePredatado(6), lChequePredatado(7), lChequePredatado(8))
End Sub
Private Sub ImpPagar(ByVal pGeral As Boolean)
    Dim x_empresa As Integer
    
    For x_empresa = 1 To 7
        lPagar(x_empresa) = MovContaPagar.TotalEntreDatas(x_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text))
        lPagar(8) = lPagar(8) + lPagar(x_empresa)
    Next
'    With tbl_contas_pagar
'        If .RecordCount > 0 Then
'            .Index = "id_data_vencimento"
'            For x_empresa = 1 To 7
'                .Seek ">=", lCodigoEmpresa(x_empresa), CDate("01/01/1900"), " ", 0
'                If Not .NoMatch Then
'                    Do Until .EOF
'                        If !Empresa <> lCodigoEmpresa(x_empresa) Then
'                            Exit Do
'                        End If
'                        If pGeral = True Or (pGeral = False And !Data_Vencimento >= CDate(msk_data_i) And !Data_Vencimento <= CDate(msk_data_f)) Then
'                            lPagar(x_empresa) = lPagar(x_empresa) + !Valor
'                            lPagar(8) = lPagar(8) + !Valor
'                        End If
'                        .MoveNext
'                    Loop
'                End If
'            Next
'        End If
'    End With
    Call ImpDet("CONTAS À PAGAR      ", lPagar(1), lPagar(2), lPagar(3), lPagar(4), lPagar(5), lPagar(6), lPagar(7), lPagar(8))
End Sub
Private Sub ImpSubTotal(ByVal pGeral As Boolean)
    Dim i As Integer
    BioImprime "@Printer.Print " & "+----------------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+--------------+"
    BioImprime "@@Printer.FontBold = True"
    Call ImpDet("** SUB TOTAL        ", lSubTotal(1), lSubTotal(2), lSubTotal(3), lSubTotal(4), lSubTotal(5), lSubTotal(6), lSubTotal(7), lSubTotal(8))
    BioImprime "@@Printer.FontBold = False"
    For i = 1 To 8
        lSubTotal(i) = 0
    Next
    BioImprime "@Printer.Print " & "+----------------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+--------------+"
End Sub
Private Sub ImpTotal(ByVal pGeral As Boolean)
    BioImprime "@@Printer.FontBold = True"
    Call ImpDet("*** TOTAL GERAL     ", lTotal(1), lTotal(2), lTotal(3), lTotal(4), lTotal(5), lTotal(6), lTotal(7), lTotal(8))
    BioImprime "@@Printer.FontBold = False"
End Sub
Private Sub ImpRodape()
    BioImprime "@Printer.Print " & "+--- Cerrado Inormática. ------------+-------------+-------------+-------------+-------------+-------------+-------------+--------------+"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpCombustivel(ByVal pGeral As Boolean)
    Dim x_empresa As Integer
    Dim xTipoCombustivel As String
    Dim xValor As Currency
    
    For x_empresa = 1 To 7
        xValor = 0
        xTipoCombustivel = "A "
        If BombaCombustivel.LocalizarTipoCombustivel(x_empresa, xTipoCombustivel) Then
            xValor = BombaCombustivel.PrecoVenda
        End If
        lCombustivelA(x_empresa) = MedicaoCombustivel.TotalMedidaCombustivel(x_empresa, CDate(msk_data.Text), xTipoCombustivel, 0) * xValor
        lCombustivelA(8) = lCombustivelA(8) + lCombustivelA(x_empresa)
    Next
    For x_empresa = 1 To 7
        xValor = 0
        xTipoCombustivel = "AA"
        If BombaCombustivel.LocalizarTipoCombustivel(x_empresa, xTipoCombustivel) Then
            xValor = BombaCombustivel.PrecoVenda
        End If
        lCombustivelAA(x_empresa) = MedicaoCombustivel.TotalMedidaCombustivel(x_empresa, CDate(msk_data.Text), xTipoCombustivel, 0) * xValor
        lCombustivelAA(8) = lCombustivelAA(8) + lCombustivelAA(x_empresa)
    Next
    For x_empresa = 1 To 7
        xValor = 0
        xTipoCombustivel = "D "
        If BombaCombustivel.LocalizarTipoCombustivel(x_empresa, xTipoCombustivel) Then
            xValor = BombaCombustivel.PrecoVenda
        End If
        lCombustivelD(x_empresa) = MedicaoCombustivel.TotalMedidaCombustivel(x_empresa, CDate(msk_data.Text), xTipoCombustivel, 0) * xValor
        lCombustivelD(8) = lCombustivelD(8) + lCombustivelD(x_empresa)
    Next
    For x_empresa = 1 To 7
        xValor = 0
        xTipoCombustivel = "DA"
        If BombaCombustivel.LocalizarTipoCombustivel(x_empresa, xTipoCombustivel) Then
            xValor = BombaCombustivel.PrecoVenda
        End If
        lCombustivelDA(x_empresa) = MedicaoCombustivel.TotalMedidaCombustivel(x_empresa, CDate(msk_data.Text), xTipoCombustivel, 0) * xValor
        lCombustivelDA(8) = lCombustivelDA(8) + lCombustivelDA(x_empresa)
    Next
    For x_empresa = 1 To 7
        xValor = 0
        xTipoCombustivel = "G "
        If BombaCombustivel.LocalizarTipoCombustivel(x_empresa, xTipoCombustivel) Then
            xValor = BombaCombustivel.PrecoVenda
        End If
        lCombustivelG(x_empresa) = MedicaoCombustivel.TotalMedidaCombustivel(x_empresa, CDate(msk_data.Text), xTipoCombustivel, 0) * xValor
        lCombustivelG(8) = lCombustivelG(8) + lCombustivelG(x_empresa)
    Next
    For x_empresa = 1 To 7
        xValor = 0
        xTipoCombustivel = "GA"
        If BombaCombustivel.LocalizarTipoCombustivel(x_empresa, xTipoCombustivel) Then
            xValor = BombaCombustivel.PrecoVenda
        End If
        lCombustivelGA(x_empresa) = MedicaoCombustivel.TotalMedidaCombustivel(x_empresa, CDate(msk_data.Text), xTipoCombustivel, 0) * xValor
        lCombustivelGA(8) = lCombustivelGA(8) + lCombustivelGA(x_empresa)
    Next
    
'    With tbl_combustivel
'        If .RecordCount > 0 Then
'            .Index = "id_codigo"
'            tbl_bomba.Index = "id_combustivel"
'            For x_empresa = 1 To 7
'                .Seek ">=", lCodigoEmpresa(x_empresa), 0
'                If Not .NoMatch Then
'                    Do Until .EOF
'                        If !Empresa <> lCodigoEmpresa(x_empresa) Then
'                            Exit Do
'                        End If
'                        x_valor = 0
'                        If ![Quantidade em Estoque] > 0 Then
'                            tbl_bomba.Seek ">=", !Empresa, !Codigo
'                            If Not tbl_bomba.NoMatch Then
'                                If tbl_bomba!Empresa = !Empresa Then
'                                    x_valor = tbl_bomba![Preco de Custo] * ![Quantidade em Estoque]
'                                End If
'                            End If
'                        End If
'                        If Trim(!Codigo) = "A" Then
'                            lCombustivelA(x_empresa) = lCombustivelA(x_empresa) + x_valor
'                            lCombustivelA(8) = lCombustivelA(8) + x_valor
'                        ElseIf Trim(!Codigo) = "AA" Then
'                            lCombustivelAA(x_empresa) = lCombustivelAA(x_empresa) + x_valor
'                            lCombustivelAA(8) = lCombustivelAA(8) + x_valor
'                        ElseIf Trim(!Codigo) = "D" Then
'                            lCombustivelD(x_empresa) = lCombustivelD(x_empresa) + x_valor
'                            lCombustivelD(8) = lCombustivelD(8) + x_valor
'                        ElseIf Trim(!Codigo) = "DA" Then
'                            lCombustivelDA(x_empresa) = lCombustivelDA(x_empresa) + x_valor
'                            lCombustivelDA(8) = lCombustivelDA(8) + x_valor
'                        ElseIf Trim(!Codigo) = "G" Then
'                            lCombustivelG(x_empresa) = lCombustivelG(x_empresa) + x_valor
'                            lCombustivelG(8) = lCombustivelG(8) + x_valor
'                        ElseIf Trim(!Codigo) = "GA" Then
'                            lCombustivelGA(x_empresa) = lCombustivelGA(x_empresa) + x_valor
'                            lCombustivelGA(8) = lCombustivelGA(8) + x_valor
'                        End If
'                        .MoveNext
'                    Loop
'                End If
'            Next
'        End If
'    End With
    If lCombustivelA(8) > 0 Then
        Call ImpDet("ESTOQUE ALCOOL      ", lCombustivelA(1), lCombustivelA(2), lCombustivelA(3), lCombustivelA(4), lCombustivelA(5), lCombustivelA(6), lCombustivelA(7), lCombustivelA(8))
    End If
    If lCombustivelAA(8) > 0 Then
        Call ImpDet("ESTOQUE ALCOOL AD.  ", lCombustivelAA(1), lCombustivelAA(2), lCombustivelAA(3), lCombustivelAA(4), lCombustivelAA(5), lCombustivelAA(6), lCombustivelAA(7), lCombustivelAA(8))
    End If
    If lCombustivelD(8) > 0 Then
        Call ImpDet("ESTOQUE DIESEL      ", lCombustivelD(1), lCombustivelD(2), lCombustivelD(3), lCombustivelD(4), lCombustivelD(5), lCombustivelD(6), lCombustivelD(7), lCombustivelD(8))
    End If
    If lCombustivelDA(8) > 0 Then
        Call ImpDet("ESTOQUE DIESEL AD.  ", lCombustivelDA(1), lCombustivelDA(2), lCombustivelDA(3), lCombustivelDA(4), lCombustivelDA(5), lCombustivelDA(6), lCombustivelDA(7), lCombustivelDA(8))
    End If
    If lCombustivelG(8) > 0 Then
        Call ImpDet("ESTOQUE GASOLINA    ", lCombustivelG(1), lCombustivelG(2), lCombustivelG(3), lCombustivelG(4), lCombustivelG(5), lCombustivelG(6), lCombustivelG(7), lCombustivelG(8))
    End If
    If lCombustivelGA(8) > 0 Then
        Call ImpDet("ESTOQUE GASOLINA AD.", lCombustivelGA(1), lCombustivelGA(2), lCombustivelGA(3), lCombustivelGA(4), lCombustivelGA(5), lCombustivelGA(6), lCombustivelGA(7), lCombustivelGA(8))
    End If
End Sub
Private Sub ImpDet(x_historico As String, x_valor_2 As Currency, x_valor_3 As Currency, x_valor_4 As Currency, x_valor_6 As Currency, x_valor_9 As Currency, x_valor_10 As Currency, x_valor_11 As Currency, x_valor_12 As Currency)
    Dim x_linha As String
    Dim i As Integer
    x_linha = Space(137)
    Mid(x_linha, 1, 1) = "|"
    Mid(x_linha, 3, 20) = x_historico
    Mid(x_linha, 24, 1) = "|"
    i = Len(Format(x_valor_2, "#####,##0.00"))
    Mid(x_linha, 25 + 12 - i, i) = Format(x_valor_2, "#####,##0.00")
    Mid(x_linha, 38, 1) = "|"
    i = Len(Format(x_valor_3, "#####,##0.00"))
    Mid(x_linha, 39 + 12 - i, i) = Format(x_valor_3, "#####,##0.00")
    Mid(x_linha, 52, 1) = "|"
    i = Len(Format(x_valor_4, "#####,##0.00"))
    Mid(x_linha, 53 + 12 - i, i) = Format(x_valor_4, "#####,##0.00")
    Mid(x_linha, 66, 1) = "|"
    i = Len(Format(x_valor_6, "#####,##0.00"))
    Mid(x_linha, 67 + 12 - i, i) = Format(x_valor_6, "#####,##0.00")
    Mid(x_linha, 80, 1) = "|"
    i = Len(Format(x_valor_9, "#####,##0.00"))
    Mid(x_linha, 81 + 12 - i, i) = Format(x_valor_9, "#####,##0.00")
    Mid(x_linha, 94, 1) = "|"
    i = Len(Format(x_valor_10, "#####,##0.00"))
    Mid(x_linha, 95 + 12 - i, i) = Format(x_valor_10, "#####,##0.00")
    Mid(x_linha, 108, 1) = "|"
    i = Len(Format(x_valor_11, "#####,##0.00"))
    Mid(x_linha, 109 + 12 - i, i) = Format(x_valor_11, "#####,##0.00")
    Mid(x_linha, 122, 1) = "|"
    i = Len(Format(x_valor_12, "#####,##0.00"))
    Mid(x_linha, 124 + 12 - i, i) = Format(x_valor_12, "#####,##0.00")
    Mid(x_linha, 137, 1) = "|"
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpCab()
    Dim x_linha As String
    Dim i As Integer
    Dim i2 As Integer
    Dim i3 As Currency
    Dim x_empresa As Integer
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
    x_linha = "|                                                                  Página, " & Format(lPagina, "000") & " |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    'g_string = ReadINI("GRUPO DE EMPRESAS", "Nome do Grupo", gArquivoIni)
    'i = Len(g_string)
    'Mid(x_linha, 3, i) = g_string
    'g_string = ""
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| ANALISE GERAL DOS POSTOS                                        , __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(x_linha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(x_linha, 69, 10) = msk_data
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| REFERENTE AO PERIODO DE.: __/__/____ A __/__/____                            |"
    Mid(x_linha, 29, 10) = msk_data_i
    Mid(x_linha, 42, 10) = msk_data_f
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & "+----------------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+--------------+"
    x_linha = "| HISTÓRICO            |             |             |             |             |             |             |             | TOTAL  GERAL |"
    
    For x_empresa = 1 To 8
        i2 = Len(Trim(lNomeEmpresa(x_empresa)))
        i3 = (13 * x_empresa + x_empresa + 11) + ((13 - i2) / 2)
        If Mid(Format(i3, "000.0"), 5, 1) <> "0" Then
            i = Val(i3) + 1
        Else
            i = Val(i3)
        End If
        Mid(x_linha, i, i2) = Trim(lNomeEmpresa(x_empresa))
    Next
    'For x_empresa = 1 To 8
    '    i2 = Len(Trim(lNomeEmpresa(x_empresa)))
    '    i3 = (13 * x_empresa + x_empresa + 11) + ((13 - i2) / 2)
    '    If Mid(Format(i3, "000.0"), 5, 1) <> "0" Then
    '        i = Val(i3) + 1
    '    Else
    '        i = Val(i3)
    '    End If
    '    Mid(x_linha, i, i2) = Trim(lNomeEmpresa(x_empresa))
    'Next
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@Printer.Print " & "+----------------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+--------------+"
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cmd_imprimir.SetFocus
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
    cmd_imprimir.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cmd_imprimir.SetFocus
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
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data de emissão.", 64, "Atenção!"
        msk_data.SetFocus
    ElseIf Not IsDate(msk_data_i) Then
        MsgBox "Informe a data inicial.", 64, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f) Then
        MsgBox "Informe a data final.", 64, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f) < CDate(msk_data_i) Then
        MsgBox "Data final deve ser maior ou igual a " & CDate(msk_data_i) & ".", 64, "Atenção!"
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
        msk_data_i.Text = Format(CDate("01/01/1900"), "dd/mm/yyyy")
        msk_data_f.Text = Format(CDate("31/12/2500"), "dd/mm/yyyy")
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
    
    MedicaoCombustivel.NomeTabela = "MedicaoCombustivelLMC"
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
        cmd_imprimir.SetFocus
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
