VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form emissao_cupom_complementar_conv 
   Caption         =   "Emissão do Cupom Complementar (Conveniência)"
   ClientHeight    =   2295
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "emissao_cupom_complementar_conv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "emissao_cupom_complementar_conv.frx":030A
   ScaleHeight     =   2295
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "emissao_cupom_complementar_conv.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Visualiza inventário de produtos."
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "emissao_cupom_complementar_conv.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprime inventário de produtos."
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "emissao_cupom_complementar_conv.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1320
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "emissao_cupom_complementar_conv.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "emissao_cupom_complementar_conv.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "emissao_cupom_complementar_conv.frx":6CBA
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
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   3840
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
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
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
Attribute VB_Name = "emissao_cupom_complementar_conv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lSQL As String

'Início de variáveis padrão para relatório
Dim lLinha As Integer
Dim lPagina As Integer
Dim lLocal As Integer
Dim lNomeArquivo As String
Dim lNomeArquivoTXT As String
'Fim de variáveis padrão para relatório
Dim BemaRetorno As Integer

Dim lImpBematech As Boolean
Dim lImpSchalter As Boolean
Dim lImpMecaf As Boolean

Dim lPeriodo As Integer
Dim lRSCriado As Boolean


Dim lNumeroCupom As Long
Dim lOrdemCupom As Integer
Dim lDataCupom As Date
Dim lHoraCupom As Date
Dim l_flag_cupom_fiscal As String
Dim lQtdCupom As Currency
Dim lValorCupom As Currency
Dim lQtdConveniencia As Currency
Dim lValorConveniencia As Currency
Dim lSerieECF As String
Dim lCodigoEcf As Integer

Private Aliquota As New cAliquota
Private ECF As New cEcf
Private LiberacaoDigitacao As New cLiberacaoDigitacao
Private MovCupomFiscal As New cMovimentoCupomFiscal
Private MovCupomFiscalItem As New cMovimentoCupomFiscalItem
Private SubGrupo As New cSubGrupo

Dim rsCupomComplementar As New adodb.Recordset
Dim rstVendaConveniencia As New adodb.Recordset
Dim rstVendaCupom As New adodb.Recordset
Private Sub AtualizaConstantes()
    Dim dados As String
    lPeriodo = 5
    If LiberacaoDigitacao.LocalizarCodigo(g_empresa, 3) Then
        lPeriodo = LiberacaoDigitacao.PeriodoInicial
    End If
    lImpBematech = False
    lImpSchalter = False
    lImpMecaf = False
    lSerieECF = dados = ReadINI("CUPOM FISCAL", "Serie ECF", gArquivoIni)
    lCodigoEcf = 1
    If ECF.LocalizarNumeroSerie(g_empresa, lSerieECF) Then
        lCodigoEcf = ECF.Codigo
    End If
    dados = ReadINI("CUPOM FISCAL", "Impressora Fiscal", gArquivoIni)
    If dados = "BEMATECH" Then
        lImpBematech = True
    ElseIf dados = "SCHALTER" Then
        lImpSchalter = True
    ElseIf dados = "MECAF" Then
        lImpMecaf = True
    End If
End Sub
Private Function AtualizaTabelaCupomFiscal(xNumeroCupom As Long, xOrdem As Integer, xData As Date, xHora As Date, xCodigoProduto As Long, xValorUnitario As Currency, xQuantidade As Currency, xValorTotal As Currency, xCodigoAliquota As Integer, xLinhaArquivo As String) As Boolean
    On Error GoTo FileError
    
    AtualizaTabelaCupomFiscal = False
    MovCupomFiscal.Empresa = g_empresa
    MovCupomFiscal.NumeroCupom = xNumeroCupom
    MovCupomFiscal.Ordem = xOrdem
    MovCupomFiscal.Data = xData
    MovCupomFiscal.Hora = xHora
    MovCupomFiscal.DataCupom = xData
    MovCupomFiscal.Periodo = CStr(lPeriodo)
    MovCupomFiscal.TipoMovimento = 1
    MovCupomFiscal.CodigoCliente = 0
    MovCupomFiscal.CodigoConveniado = 0
    MovCupomFiscal.CodigoProduto = xCodigoProduto
    MovCupomFiscal.ValorUnitario = xValorUnitario
    MovCupomFiscal.Quantidade = xQuantidade
    MovCupomFiscal.ValorTotal = xValorTotal
    MovCupomFiscal.FormaPagamento = 1
    MovCupomFiscal.ValorRecebido = xValorTotal
    MovCupomFiscal.NumeroCheque = ""
    MovCupomFiscal.Telefone = ""
    MovCupomFiscal.operador = 1
    MovCupomFiscal.CupomCancelado = False
    MovCupomFiscal.ItemCancelado = False
    MovCupomFiscal.CodigoAliquota = xCodigoAliquota
    MovCupomFiscal.ValorDesconto = 0
    MovCupomFiscal.Nome = ".,."
    MovCupomFiscal.CPFCNPJ = ""
    MovCupomFiscal.ValorDescontoEmbutido = 0
    If MovCupomFiscal.Incluir Then
        AtualizaTabelaCupomFiscal = True
    Else
        Call CriaLogCupom(Time & " - Emissão do Cupom Complementar Conveniência: Não foi possível gravar este cupom - xNumeroCupom=" & xNumeroCupom & " - xOrdem=" & xOrdem & " - xData=" & xData)
        MsgBox "Não foi possível gravar este Cupom Complementar!", vbInformation, "Erro de Integridade!"
    End If
    Exit Function
    
FileError:
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar Conveniência: Erro na Rotina: AtualizaTabelaCupomFiscal - xNumeroCupom=" & xNumeroCupom & " - xOrdem=" & xOrdem & " - xData=" & xData & " - Error=" & Error)
    MsgBox "Não foi possível gravar este Cupom Complementar!", vbInformation, "Erro de Integridade!"
    Exit Function
End Function
Private Sub BuscaNumeroCupom()
    Dim xString As String
    Dim NumeroArquivo As Integer
    Dim xData As String
    Dim xHora As String
    
    On Error GoTo FileError
    
    If Not Testa_ImpressoraCF Then
        NumeroArquivo = 99999
    End If
    If l_flag_cupom_fiscal = "F" Then
        l_flag_cupom_fiscal = "A"
        
        'busca numero do cupom da impressora fiscal
        xString = Space(6)
        BemaRetorno = Bematech_FI_NumeroCupom(xString)
        If BemaRetorno <> 1 Then
            Call AnalizaRetornoBematech(BemaRetorno)
        End If
        lNumeroCupom = CLng(xString) + 1
    End If
    
    'busca data/hora da impressora fiscal
    xData = Space(6)
    xHora = Space(6)
    BemaRetorno = Bematech_FI_DataHoraImpressora(xData, xHora)
    lDataCupom = CDate(Mid(xData, 1, 2) & "/" & Mid(xData, 3, 2) & "/20" & Mid(xData, 5, 2))
    lHoraCupom = Format(Mid(xHora, 1, 2), "00") & ":" & Format(Mid(xHora, 3, 2), "00") & ":" & Format(Mid(xHora, 5, 2), "00")
    Exit Sub

FileError:
    MsgBox "Não foi possível criar o novo cupom fiscal.", vbCritical, "BuscaNumeroCupom"
    Exit Sub
End Sub
Private Sub CriaRsCupomComplementar()
    With rsCupomComplementar
        If lRSCriado Then
            If .RecordCount > 0 Then
                .MoveFirst
                Do Until .EOF
                    .Delete
                    .MoveNext
                Loop
            End If
        Else
            .CursorLocation = adUseClient
            .Fields.Append "Codigo", adVarChar, 6
            .Fields.Append "Nome", adVarChar, 40
            .Fields.Append "Unidade", adVarChar, 2
            .Fields.Append "Aliquota", adVarChar, 2
            .Fields.Append "QuantidadeConveniencia", adVarChar, 12
            .Fields.Append "TotalConveniencia", adVarChar, 12
            .Fields.Append "QuantidadeCupom", adVarChar, 12
            .Fields.Append "TotalCupom", adVarChar, 12
            .Open
            lRSCriado = True
        End If
    End With
End Sub
Private Sub AtivaBotoes(xAtiva As Boolean)
    cmd_visualizar.Enabled = xAtiva
    cmd_imprimir.Enabled = xAtiva
    cmd_sair.Enabled = xAtiva
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Aliquota = Nothing
    Set ECF = Nothing
    Set LiberacaoDigitacao = Nothing
    Set MovCupomFiscal = Nothing
    Set MovCupomFiscalItem = Nothing
    Set SubGrupo = Nothing
    rsCupomComplementar.Close
    Set rsCupomComplementar = Nothing
End Sub
Private Sub GravaRsCupomComplementar()
    Do Until rstVendaConveniencia.EOF
        If SubGrupo.LocalizarCodigo(rstVendaConveniencia("Codigo do SubGrupo").Value) Then
            If SubGrupo.Nome Like "*...*" Then
            Else
                rsCupomComplementar.AddNew
                rsCupomComplementar!Codigo = Format(rstVendaConveniencia("Codigo do Produto").Value, "000000")
                rsCupomComplementar!Nome = rstVendaConveniencia("Nome").Value
                rsCupomComplementar!Unidade = rstVendaConveniencia("Unidade").Value
                rsCupomComplementar!Aliquota = rstVendaConveniencia("Aliquota").Value
                rsCupomComplementar!QuantidadeConveniencia = Format(rstVendaConveniencia("TQuantidade").Value, "000000000.00")
                rsCupomComplementar!TotalConveniencia = Format(rstVendaConveniencia("Total").Value, "000000000.00")
                rsCupomComplementar!QuantidadeCupom = "000000000000"
                rsCupomComplementar!TotalCupom = "000000000000"
                rsCupomComplementar.Update
            End If
        End If
        rstVendaConveniencia.MoveNext
    Loop
End Sub
Private Sub RegravaRsCupomComplementar()
    Do Until rstVendaCupom.EOF
        
        rsCupomComplementar.Sort = "Codigo"
        rsCupomComplementar.Find "Codigo='" & Format(rstVendaCupom("Codigo do Produto").Value, "000000") & "'"
        
        If rsCupomComplementar.EOF Then
        Else
            rsCupomComplementar!QuantidadeCupom = Format(rstVendaCupom("TQuantidade").Value, "000000000.00")
            rsCupomComplementar!TotalCupom = Format(rstVendaCupom("Total").Value, "000000000.00")
            rsCupomComplementar.Update
        End If
        rstVendaCupom.MoveNext
    Loop
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lQtdCupom = 0
    lValorCupom = 0
    lQtdConveniencia = 0
    lValorConveniencia = 0
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    CriaRsCupomComplementar
    
    'Verifica Movimento_Venda_Conveniencia
    lSQL = ""
    lSQL = lSQL & "SELECT Movimento_Venda_Conveniencia.[Codigo do Produto], Produto.Nome as Nome, Produto.Unidade, Produto.[Codigo da Aliquota] AS Aliquota, Produto.[Codigo do SubGrupo], SUM(Movimento_Venda_Conveniencia.Quantidade) AS TQuantidade, SUM(Movimento_Venda_Conveniencia.[Valor Total]) AS Total"
    lSQL = lSQL & "  FROM Movimento_Venda_Conveniencia, Produto"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND Data <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & "   AND [Item Cancelado] = False"
    lSQL = lSQL & "   AND [Cupom Cancelado] = False"
    lSQL = lSQL & "   AND Produto.Codigo = Movimento_Venda_Conveniencia.[Codigo do Produto]"
    lSQL = lSQL & " GROUP BY [Codigo do Produto], Nome, Unidade, Produto.[Codigo da Aliquota], Produto.[Codigo do SubGrupo]"
    rstVendaConveniencia.Open lSQL, cnnSGP, adOpenForwardOnly, adLockReadOnly
    If Not rstVendaConveniencia.EOF Then
        GravaRsCupomComplementar
        
        'Verifica Movimento_Cupom_Fiscal
        lSQL = ""
        lSQL = lSQL & "SELECT Movimento_Cupom_Fiscal.[Codigo do Produto], SUM(Movimento_Cupom_Fiscal.Quantidade) AS TQuantidade, SUM(Movimento_Cupom_Fiscal.[Valor Total]) AS Total"
        lSQL = lSQL & "  FROM Movimento_Cupom_Fiscal"
        lSQL = lSQL & " WHERE Empresa = " & g_empresa
        lSQL = lSQL & "   AND Data >= " & preparaData(msk_data_i.Text)
        lSQL = lSQL & "   AND Data <= " & preparaData(msk_data_f.Text)
        lSQL = lSQL & " GROUP BY [Codigo do Produto]"
        rstVendaCupom.Open lSQL, cnnSGP, adOpenForwardOnly, adLockReadOnly
        If Not rstVendaCupom.EOF Then
            RegravaRsCupomComplementar
        End If
        
        ImpDados
        rstVendaCupom.Close
    Else
        Call CriaLogCupom(Time & " - Emissão do Cupom Complementar de Conveniência: Foi Informado ao Usuário que Não Existia Venda de Conveniência Digitada no Período Informado.")
        MsgBox "Não existe venda de conveniência digitada no período informado.", vbInformation, "Sem Movimento!"
    End If
    rstVendaConveniencia.Close
    Set rstVendaConveniencia = Nothing
    Call AtivaBotoes(True)
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    Dim xPrecoUnitario As Currency
    
    If rsCupomComplementar.RecordCount > 0 Then
        rsCupomComplementar.Sort = "Nome"
        rsCupomComplementar.MoveFirst
    End If
    Do Until rsCupomComplementar.EOF
        xPrecoUnitario = Format(rsCupomComplementar("TotalConveniencia").Value / rsCupomComplementar("QuantidadeConveniencia").Value, "0000000000.00")
        Call ImpDet(CLng(rsCupomComplementar("Codigo").Value), rsCupomComplementar("Nome").Value, rsCupomComplementar("Unidade").Value, fValidaValor(rsCupomComplementar("QuantidadeCupom").Value), fValidaValor(rsCupomComplementar("TotalCupom").Value), fValidaValor(rsCupomComplementar("QuantidadeConveniencia").Value), fValidaValor(rsCupomComplementar("TotalConveniencia").Value), Val(rsCupomComplementar("Aliquota").Value), xPrecoUnitario)
        rsCupomComplementar.MoveNext
    Loop
    
    If lPagina > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        If lLocal = 1 Then
            'If (MsgBox("Após a emissão do cupom complementar será impresso a REDUÇÃO Z." & Chr(13) & "E não será mais aceito a emissão de cupom fiscal nesta data." & Chr(13) & Chr(13) & "Deseja realmente imprimir o cupom complementar?", vbQuestion + vbYesNo + vbDefaultButton2, "Emissão do Cupom Complementar")) = 6 Then
                ImpCupomComplementar
                Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: O Usuário foi Informado da Conclusão da Impressão")
                MsgBox "Impressão do cupom complementar concluída", vbInformation, "Impressão Concluída!"
            'End If
        Else
            g_string = lLocal & lNomeArquivo & "|@|Emissão do Cupom Complementar|@|"
            frm_preview.Show 1
        End If
    End If
End Sub
Private Sub ImpDet(pCodigo As Long, pNome As String, pUnidade As String, pQtdCupom As Currency, pValorCupom As Currency, pQtdConveniencia As Currency, pValorConveniencia As Currency, pCodigoAliquota As Integer, pPrecoVenda As Currency)
    Dim xLinha As String
    Dim i As Integer
    Dim xQtd As Currency
    Dim xValor As Currency
    Dim xNomeProduto As String
    Dim xUnidade As String
    
    If lPagina = 0 Then
        ImpCab
    End If
    If lLinha >= 60 Then
        xLinha = "+------+-------------------------------------------+---+----------+---------------+----------+---------------+----------+---------------+"
        Mid(xLinha, 12, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    xLinha = "|      |                                           |   |          |               |          |               |          |               |"
    i = Len(Format(pCodigo, "#,000"))
    Mid(xLinha, 2 + 5 - i, i) = Format(pCodigo, "#,000")
    Mid(xLinha, 10, 40) = pNome
    Mid(xLinha, 53, 3) = pUnidade
    i = Len(Format(pQtdCupom, "###,##0.00"))
    Mid(xLinha, 57 + 10 - i, i) = Format(pQtdCupom, "###,##0.00")
    i = Len(Format(pValorCupom, "###,###,##0.00"))
    Mid(xLinha, 69 + 14 - i, i) = Format(pValorCupom, "###,###,##0.00")
    i = Len(Format(pQtdConveniencia, "###,###,##0.00"))
    Mid(xLinha, 80 + 14 - i, i) = Format(pQtdConveniencia, "###,###,##0.00")
    i = Len(Format(pValorConveniencia, "###,###,##0.00"))
    Mid(xLinha, 96 + 14 - i, i) = Format(pValorConveniencia, "###,###,##0.00")
    xQtd = Format(pQtdConveniencia - pQtdCupom, "000,000,000.00")
    i = Len(Format(xQtd, "###,###,##0.00"))
    Mid(xLinha, 107 + 14 - i, i) = Format(xQtd, "###,###,##0.00")
    xValor = Format(xQtd * pPrecoVenda, "000,000,000.00")
    i = Len(Format(xValor, "###,###,##0.00"))
    Mid(xLinha, 123 + 14 - i, i) = Format(xValor, "###,###,##0.00")
    lQtdCupom = lQtdCupom + pQtdCupom
    lValorCupom = lValorCupom + pValorCupom
    lQtdConveniencia = lQtdConveniencia + pQtdConveniencia
    lValorConveniencia = lValorConveniencia + pValorConveniencia
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
    If lLocal = 1 Then
        xNomeProduto = Space(40)
        xUnidade = Space(2)
        Mid(xNomeProduto, 1, 40) = Mid(pNome, 1, 40)
        Mid(xUnidade, 1, 2) = pUnidade
        If (pQtdConveniencia - pQtdCupom) > 0 Then
            xLinha = Format(pCodigo, "0000")
            xLinha = xLinha & Mid(xNomeProduto, 1, 40)
            xLinha = xLinha & Mid(xUnidade, 1, 2)
            xLinha = xLinha & Format(pCodigoAliquota, "00")
            xLinha = xLinha & Format(xQtd, "0000000000.00")
            xLinha = xLinha & Format(pPrecoVenda, "0000000000.0000")
            xLinha = xLinha & Format(xValor, "0000000000.00")
            'Print #3, xLinha
            gArquivoTXT.WriteLine (xLinha)
        End If
    End If
End Sub
Private Sub ImpTotal()
'    Dim y_local As Single
    Dim x_linha As String
    Dim i As Integer
    
    If lLocal = 1 Then
        'Print #3, "FIM"
        gArquivoTXT.WriteLine ("FIM")
    End If
    
    
    x_linha = "+------+-------------------------------------------+---+----------+---------------+----------+---------------+----------+---------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                               *** TOTAL DO RELATORIO |          |               |          |               |          |               |"
    i = Len(Format(lQtdCupom, "###,##0.00"))
    Mid(x_linha, 57 + 10 - i, i) = Format(lQtdCupom, "###,##0.00")
    i = Len(Format(lValorCupom, "###,###,##0.00"))
    Mid(x_linha, 69 + 14 - i, i) = Format(lValorCupom, "###,###,##0.00")
    i = Len(Format(lQtdConveniencia, "###,###,##0.00"))
    Mid(x_linha, 80 + 14 - i, i) = Format(lQtdConveniencia, "###,###,##0.00")
    i = Len(Format(lValorConveniencia, "###,###,##0.00"))
    Mid(x_linha, 96 + 14 - i, i) = Format(lValorConveniencia, "###,###,##0.00")
    i = Len(Format(lQtdConveniencia - lQtdCupom, "###,###,##0.00"))
    Mid(x_linha, 107 + 14 - i, i) = Format(lQtdConveniencia - lQtdCupom, "###,###,##0.00")
    i = Len(Format(lValorConveniencia - lValorCupom, "###,###,##0.00"))
    Mid(x_linha, 123 + 14 - i, i) = Format(lValorConveniencia - lValorCupom, "###,###,##0.00")
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
    x_linha = "+------------------------------------------------------+----------+---------------+----------+---------------+----------+---------------+"
    Mid(x_linha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
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
    BioImprime "@@Printer.Print " & Chr(34) & " " & Chr(34)
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.CurrentY = 0"
    x_linha = "+---------------------------------------------------------------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|                                                                                                                           Página: ___ |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    Mid(x_linha, 133, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| CUPOM COMPLEMENTAR DE CONVENIENCIA DO PERIODO....: __/__/____ A __/__/____.                                        Cidade, __/__/____ |"
    Mid(x_linha, 54, 10) = msk_data_i.Text
    Mid(x_linha, 67, 10) = msk_data_f.Text
    i = Len(g_cidade_empresa)
    Mid(x_linha, 94 + 30 - i, i) = g_cidade_empresa
    Mid(x_linha, 126, 10) = msk_data.Text
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    x_linha = "+------+-------------------------------------------+---+--------------------------+--------------------------+--------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|CODIGO|                                           |   | C U P O M    F I S C A L |        V E N D A S       |   CUPOM   COMPLEMENTAR   |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|  DO  | DISCRIMINAÇÃO DOS PRODUTOS                |UN.+----------+---------------+----------+---------------+----------+---------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| PROD.|                                           |   |QUANTIDADE|   V A L O R   |QUANTIDADE|   V A L O R   |QUANTIDADE|   V A L O R   |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|      |                                           |   |          |               |          |               |          |               |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "+------+-------------------------------------------+---+----------+---------------+----------+---------------+----------+---------------+"
    BioImprime "@Printer.Print " & x_linha
End Sub
Private Sub ImpCupomComplementar()
    Dim x_linha As String
    Dim xString As String
    Dim xString2 As String
    Dim xDescricao As String
    Dim i As Integer
    
    Dim CodigoProduto As String
    Dim NomeProduto As String
    Dim xAliquota As String
    Dim Quantidade As String
    Dim Valor As String
    Dim ValorDesconto As String
    Dim ValorAcrescimo As String
    Dim Departamento As String
    Dim Un As String
    
    Dim x_valor_acrescimo As Currency
    Dim x_valor_desconto As Currency
    Dim xTotalECF As Currency
    Dim xValorUnitario As String * 9
    Dim xValorUnitario2 As Currency
    Dim xValorTotal As Currency
    Dim xQuantidade As String * 7
    Dim xQuantidade2 As Currency
    Dim xCodigoProduto As Long
    Dim xCodigoAliquota As Integer
    Dim xCodigoFiscal As String * 2
    
    On Error GoTo ErroImpCupomComplementar
    
    gArquivoTXT.Close
        
    'Verifica se existe Cupom para imprimir
    Set gArquivoTXT = gArqTxt.OpenTextFile(lNomeArquivoTXT, ForReading)
    x_linha = gArquivoTXT.ReadLine
    gArquivoTXT.Close
    If Mid(x_linha, 1, 3) = "FIM" Then
        Call CriaLogCupom(Time & " - Emissão do Cupom Complementar de Conveniência: Não existe ítem a ser impresso")
        Exit Sub
    End If
    
    'Abre Cupom Fiscal
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar de Conveniência: Foi Acionado a Abertura do CF")
    l_flag_cupom_fiscal = "F"
    Call BuscaNumeroCupom
    lOrdemCupom = 0
    xTotalECF = 0
    'Abre o cupom fiscal
    BemaRetorno = Bematech_FI_AbreCupom("")
    
    
    'Loop para imprimir os itens do cupom
    Set gArquivoTXT = gArqTxt.OpenTextFile(lNomeArquivoTXT, ForReading)
    Do Until gArquivoTXT.AtEndOfStream
        x_linha = gArquivoTXT.ReadLine
        If Mid(x_linha, 1, 3) = "FIM" Then
            Exit Do
        End If
        
        Call CriaLogCupom(Time & " - Emissão do Cupom Complementar de Conveniência: Foi Acionado a Emissão do C.F. da linha =" & x_linha)
        lOrdemCupom = lOrdemCupom + 1
        
        'Venda de Item com entrada de departamento,
        'Verifica se há diferença do total
        xString = Format(Format(fValidaValor(Mid(x_linha, 63, 15)) * fValidaValor(Mid(x_linha, 50, 13)), "###,##0.0000"), "###,##0.0000")
        i = Len(xString)
        xString = Mid(xString, 1, i - 2)
        x_valor_acrescimo = 0
        x_valor_desconto = 0
        If fValidaValor(Mid(x_linha, 78, 13)) > fValidaValor(xString) Then
            x_valor_acrescimo = fValidaValor(Mid(x_linha, 78, 13)) - fValidaValor(xString)
        ElseIf fValidaValor(Mid(x_linha, 78, 13)) < fValidaValor(xString) Then
            x_valor_desconto = fValidaValor(xString) - fValidaValor(Mid(x_linha, 78, 13))
        Else
        End If
        
        
        'código do produto
        xCodigoProduto = Mid(x_linha, 1, 4)
        CodigoProduto = Format(Mid(x_linha, 1, 4), "#,##0")
        'nome do produto
        NomeProduto = Mid(x_linha, 5, 40)
        'tipo de tributação
        xCodigoAliquota = Mid(x_linha, 47, 2)
        If Aliquota.LocalizarCodigo(lSerieECF, xCodigoAliquota) Then
            xAliquota = Aliquota.CodigoFiscal
        Else
            xAliquota = "II"
        End If
        'Valor Unitário
        xString = Format(Mid(x_linha, 63, 15), "000000.000")
        Valor = Mid(xString, 1, 6) + Mid(xString, 8, 3)
        xValorUnitario2 = xString
        'Quantidade
        xString = Format(Mid(x_linha, 50, 13), "0000.000")
        Quantidade = Mid(xString, 1, 4) + Mid(xString, 6, 3)
        xQuantidade2 = Format(Mid(x_linha, 50, 13), "0000.000")
        'Valor do Acréscimo
        xString = Format(x_valor_acrescimo, "00000000.00")
        ValorAcrescimo = Mid(xString, 1, 8) + Mid(xString, 10, 2)
        'Valor do Desconto
        xString = Format(x_valor_desconto, "00000000.00")
        ValorDesconto = Mid(xString, 1, 8) + Mid(xString, 10, 2)
        
        'Desconsidera Descontos ou Acréscimos
        If x_valor_acrescimo > 0 Or x_valor_desconto > 0 Then
            x_valor_acrescimo = 0
            x_valor_desconto = 0
            ValorAcrescimo = "0000000000"
            ValorDesconto = "0000000000"
        End If
        
        'Departamento
        Departamento = Format(1, "00")
        'Unidade de Medida
        Un = Mid(x_linha, 45, 2)
        
        'Imprime Item
        BemaRetorno = Bematech_FI_VendeItemDepartamento(CodigoProduto, NomeProduto, xAliquota, Valor, Quantidade, ValorAcrescimo, ValorDesconto, Departamento, Un)
        If BemaRetorno = 1 Then
            'Grava Cupom Complementar
            xValorTotal = Format(xValorUnitario2 * xQuantidade2, "0000000000.00") - x_valor_desconto + x_valor_acrescimo
            If AtualizaTabelaCupomFiscal(lNumeroCupom, lOrdemCupom, lDataCupom, lHoraCupom, xCodigoProduto, xValorUnitario2, xQuantidade2, xValorTotal, xCodigoAliquota, x_linha) Then
                xTotalECF = xTotalECF + xValorTotal
            End If
        Else
            Call AnalizaRetornoBematech(BemaRetorno)
        End If
    Loop
    
    'Finaliza ECF
    
    'Desconto para o Cupom Fiscal
    xString = Mid(Format(fValidaValor(0), "000000000000.00"), 1, 12) + Mid(Format(fValidaValor(0), "000000000000.00"), 14, 2)
    BemaRetorno = Bematech_FI_IniciaFechamentoCupom("D", "$", xString)
    
    'Efetua Forma de Pagamento
    xString = "Dinheiro        "
    xString2 = Mid(Format(xTotalECF, "000000000000.00"), 1, 12) + Mid(Format(xTotalECF, "000000000000.00"), 14, 2)
    xDescricao = ""
    BemaRetorno = Bematech_FI_EfetuaFormaPagamentoDescricaoForma(xString, xString2, xDescricao)
    
    'Fecha Cupom Fiscal
    xString = "Cerrado Informatica - (062) 8436-4444           Sistemas para Automacao Comercial               "
    BemaRetorno = Bematech_FI_TerminaFechamentoCupom(xString)
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar Conveniência: Foi finalizado com sucesso!")
    l_flag_cupom_fiscal = "F"
    
    'Altera o total do ECF
    MovCupomFiscal.ValorRecebido = xTotalECF
    If Not MovCupomFiscal.AlterarFormaPagamento(g_empresa, lCodigoEcf, lNumeroCupom, lDataCupom) Then
        Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Erro ImpCupomComplementar - Não foi possível alterar a forma de pagamento!")
    End If
    
    Exit Sub
ErroImpCupomComplementar:
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar: Erro ImpCupomComplementar - " & x_linha)
    Exit Sub
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = RetiraGString(1)
    cmd_visualizar.SetFocus
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
    cmd_visualizar.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cmd_visualizar.SetFocus
    Else
        msk_data_i = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    On Error GoTo ErroImprimir
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar de Conveniência: Foi Pedido a Impressão do Período de: " & msk_data_i.Text & " a " & msk_data_f.Text)
    
    'Cria o Arquivo CUPOM_COMPLEMENTAR.TXT
    lNomeArquivoTXT = "\VB5\SGP\DATA\CUPOM_COMPLEMENTAR.TXT"
    Set gArquivoTXT = gArqTxt.CreateTextFile(lNomeArquivoTXT, True)
    
    lLocal = 1
    If ValidaCampos Then
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Call AtivaBotoes(False)
            g_string = "imprimiu|@|"
            Relatorio
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar de Conveniência: Foi Impresso")
        End If
    End If
    gArquivoTXT.Close
    Exit Sub
ErroImprimir:
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar de Conveniência: Erro ao abrir o arquivo CUPOM_COMPLEMENTAR.TXT")
    Call AtivaBotoes(True)
    Exit Sub
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
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar de Conveniência: Foi Pedido a Visualização do Período de: " & msk_data_i.Text & " a " & msk_data_f.Text)
    lLocal = 0
    If ValidaCampos Then
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 6, "")
            Call AtivaBotoes(False)
            Relatorio
            Call CriaLogCupom(Time & " - Emissão do Cupom Complementar de Conveniência: Foi Visualizado")
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar (5): A Emissão do Cupom Complementar Foi Aberta")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        msk_data_i.Text = Format(Date, "dd/mm/yyyy")
        msk_data_f.Text = Format(Date, "dd/mm/yyyy")
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
    Call CriaLogCupom(Time & " - Emissão do Cupom Complementar de Conveniência(1): Foi Aberta.")
    CentraForm Me
    
    AtualizaConstantes
    lRSCriado = False
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
        cmd_visualizar.SetFocus
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
