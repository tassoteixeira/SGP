VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form GeraArquivoSintegra 
   Caption         =   "Gera Arquivo para o Sintegra"
   ClientHeight    =   2760
   ClientLeft      =   1170
   ClientTop       =   1065
   ClientWidth     =   6810
   Icon            =   "GeraArquivoSintegra.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2760
   ScaleWidth      =   6810
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   1800
      Picture         =   "GeraArquivoSintegra.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Confirma a geração do disquete para o sintegra."
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4200
      Picture         =   "GeraArquivoSintegra.frx":15E4
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1800
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1635
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6555
      Begin VB.Frame frm_finalidade 
         Caption         =   "Finalidade"
         Height          =   555
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   3315
         Begin VB.OptionButton opt_finalidade_normal 
            Caption         =   "Normal"
            Height          =   255
            Left            =   180
            TabIndex        =   8
            Top             =   240
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.OptionButton opt_finalidade_retificacao 
            Caption         =   "Retificação"
            Height          =   255
            Left            =   1800
            TabIndex        =   9
            Top             =   240
            Width           =   1395
         End
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "GeraArquivoSintegra.frx":28BE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2820
         Picture         =   "GeraArquivoSintegra.frx":3B98
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4800
         TabIndex        =   5
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
      Begin MSMask.MaskEdBox msk_data_i 
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
         Caption         =   "Data &Final"
         Height          =   315
         Left            =   3780
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "D&ata Inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
End
Attribute VB_Name = "GeraArquivoSintegra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lLinhaDados As String
Dim lRegistro As Long
Dim lRegistro50 As Long
Dim lRegistro53 As Long
Dim lRegistro54 As Long
Dim lRegistro60A As Long
Dim lRegistro60M As Long
Dim lRegistro70 As Long
Dim lRegistro71 As Long
Dim lRegistro75 As Long
Dim lSQL As String

Dim lCNPJ As String
Dim lIE As String
Dim lDataEmissao As String
Dim lUF As String
Dim lCodModDocFisc As String
Dim lSerie As String
Dim lSubSerie As String
Dim lNumeroNF As String
Dim lCFOP As String
Dim lEmitente As String
Dim lValorNF As String
Dim lBaseCalculo As String
Dim lValorICMS As String
Dim lIsentas As String
Dim lOutras As String
Dim lAliquotaICMS As String
Dim lNFCancelada As String
Dim lNumeroSerieECF(1 To 20) As String

Dim lEmpresaCNPJ As String
Dim lEmpresaIE As String
Dim lEmpresaUF As String

Dim lrsItemEntradaSaidaCriado As Boolean
Dim rsItemEntradaSaida As New adodb.Recordset
Dim rsEntradaCabecalho As New adodb.Recordset
Dim rsTabela As adodb.Recordset

Dim Aliquota As New CadastroDLL.cAliquota
Dim Bomba As New CadastroDLL.cBomba
Dim ECF As New CadastroDLL.cEcf
Dim Empresa As New CadastroDLL.cEmpresa
Dim Fornecedor As New CadastroDLL.cFornecedor
Dim Produto As New CadastroDLL.cProduto

Function AliquotaICMS(x_data As Date, x_impressora As Integer, x_numero As Long, x_serie As String) As Currency
'    'loop tabela Movimento_Nota_Fiscal_Saida
'    AliquotaICMS = 0
'    With tbl_movimento_nota_fiscal_saida
'        .Seek ">=", g_empresa, x_data, x_impressora, x_numero, x_serie, 0
'        If Not .NoMatch Then
'            If !Empresa = g_empresa And !Data = x_data And !Impressora = x_impressora And !numero = x_numero And !Serie = x_serie Then
'                'tbl_aliquota.Seek "=", ![Codigo da Aliquota]
'                'If Not tbl_aliquota.NoMatch Then
'                '    AliquotaICMS = tbl_aliquota![Aliquota do Imposto]
'                'End If
'                AliquotaICMS = ![Aliquota de ICMS]
'            End If
'        End If
'    End With
End Function
Function AliquotaItemICMS(ByVal pCodigoAliquota As Integer) As Currency
    AliquotaItemICMS = 0
    If Aliquota.LocalizarCodigoAliquota(pCodigoAliquota) Then
        AliquotaItemICMS = Aliquota.Aliquota
    End If
End Function
Private Sub LimpaArquivosTemporarioValidadorSintegra()
    On Error GoTo ErrorLimpaArquivos
    Kill "C:\Arquivos de programas\Validador Sintegra\NFTEMP*.*"
    Exit Sub
ErrorLimpaArquivos:
    Exit Sub
End Sub
Private Function NomeArquivo() As String
    Dim xNomeArquivo As String
    Dim xNomeEmpresa As String
    Dim xVetor As Variant
    Dim i As Integer
    
    NomeArquivo = "SINTEGRA.TXT"
    xNomeArquivo = "Sintegra_" & Mid(msk_data_i.Text, 4, 2) & "_" & Mid(msk_data_i.Text, 7, 4) & "_"
    xNomeEmpresa = ""
    xVetor = Split(Empresa.Nome, " ")
    For i = LBound(xVetor) To UBound(xVetor)
        If UCase(xVetor(i)) <> "LTDA" Then
            xVetor(i) = LCase(xVetor(i))
            Mid(xVetor(i), 1, 1) = UCase(Mid(xVetor(i), 1, 1))
            xNomeEmpresa = xNomeEmpresa & xVetor(i)
        End If
    Next
    xNomeArquivo = xNomeArquivo & xNomeEmpresa & ".TXT"
    NomeArquivo = xNomeArquivo
End Function
Private Sub PreparaBotoes()
    cmd_ok.Visible = True
    frm_dados.Enabled = True
End Sub
Private Sub GeraArquivo()
    Dim x_data_vencimento As Date
    Dim x_valor As Currency
    Dim i As Integer
    
    On Error GoTo ErrorGeraArquivo
    
    Open (NomeArquivo) For Output As #1
    lRegistro = 0
    lRegistro50 = 0
    lRegistro53 = 0
    lRegistro54 = 0
    lRegistro60A = 0
    lRegistro60M = 0
    lRegistro70 = 0
    lRegistro71 = 0
    lRegistro75 = 0
    For i = 1 To 10
        lNumeroSerieECF(i) = Space(20)
        If ECF.LocalizarCodigo(g_empresa, i) Then
            Mid(lNumeroSerieECF(i), 1, Len(ECF.NumeroSerie)) = ECF.NumeroSerie
        End If
    Next
    
    CriarsItemEntradaSaida
    GeraRsItemEntradaCombustivel
    GeraRsItemEntradaProduto
    'GeraRsItemSaida 'nesta rotina será das outras notas de saidas
    GravaArquivoRegistro10e11
    LoopRsEntradaCombustivel (50)
    'LoopRsEntradaCombustivel (53)
'    GravaArquivoRegistro50Saida
    GravaArquivoRegistro54
    
    'Aqui a contadora pediu para nao gerar mapa resumo
    'Nao concordei, mas fiz conforme pedido dela
    'LoopRsMapaResumo
    
    
    
    '70
'    LoopRsEntradaCombustivel (True)
'    GravaArquivoRegistro71
    GravaArquivoRegistro75
    GravaArquivoRegistro90
    Close #1
    MsgBox "O arquivo foi gerado com sucesso.", vbExclamation, "Fim de Geração de Arquivo!"
    Exit Sub
ErrorGeraArquivo:
    Close #1
    'If Err = 53 Then
    '    MsgBox "Arquivo não encontrado para a data informada.", vbCritical, "Erro no processamento"
    '    Exit Sub
    'Else
        MsgBox "Não foi possível processar o disquete." & Chr(10) & "Erro de número " & Err, vbCritical, "GeraArquivo"
        Exit Sub
    'End If
End Sub
Private Sub GravaArquivoRegistro10e11()
    Dim i As Integer
    Dim i2 As Integer
    'Dim x_nome_empresa As String * 35
    'Dim x_endereco As String * 76
    Dim xString As String
    
    On Error GoTo ErrorRotina
    
    lEmpresaIE = "              "
    lEmpresaCNPJ = Mid(Empresa.CGC, 1, 14)
    lEmpresaUF = Empresa.Estado
    i2 = 0
    For i = 1 To 20
        If Mid(Empresa.InscricaoEstadual, i, 1) >= "0" And Mid(Empresa.InscricaoEstadual, i, 1) <= "9" Then
            i2 = i2 + 1
            Mid(lEmpresaIE, i2, 1) = Mid(Empresa.InscricaoEstadual, i, 1)
        End If
    Next
    'Registro Tipo 10
    'Mestre do Estabelecimento
    lLinhaDados = "10"
    'Numero do CNPJ
    lLinhaDados = lLinhaDados & Empresa.CGC
    'Inscricao Estadual
    lLinhaDados = lLinhaDados & lEmpresaIE
    'Razao Social
    xString = Space(35)
    Mid(xString, 1, 35) = Empresa.Nome
    lLinhaDados = lLinhaDados & xString
    'Cidade
    xString = Space(30)
    Mid(xString, 1, 30) = Empresa.Cidade
    lLinhaDados = lLinhaDados & xString
    'UF
    lLinhaDados = lLinhaDados & UCase(Empresa.Estado)
    'FAX
    xString = Space(10)
    If Len(Empresa.Telefone) = 8 Then
        Mid(xString, 1, 10) = "62" & Mid(Empresa.Telefone, 1, 8)
    ElseIf Len(Empresa.Telefone) = 11 Then
        Mid(xString, 1, 10) = Mid(Empresa.Telefone, 2, 10)
    Else
        Mid(xString, 1, 10) = Empresa.Telefone
    End If
    lLinhaDados = lLinhaDados & xString
    'Data Inicial
    lLinhaDados = lLinhaDados & Mid(msk_data_i.Text, 7, 4) & Mid(msk_data_i.Text, 4, 2) & Mid(msk_data_i.Text, 1, 2)
    'Data Final
    lLinhaDados = lLinhaDados & Mid(msk_data_f.Text, 7, 4) & Mid(msk_data_f.Text, 4, 2) & Mid(msk_data_f.Text, 1, 2)
    'Código da identificação do Convênio
    lLinhaDados = lLinhaDados & "3" 'antes de 2003 = 1 e depois = 2
    'Código da identificação da natureza das operações informadas
    '1-Interestaduais (somente com substituição tributária)
    '2-Interestaduais (com ou sem substituição tributária)
    '3-Totalidade das informações do informante
    lLinhaDados = lLinhaDados & "3"
    'Código da finalidade do arquivo magnético
    '1=Normal  2=Retificação
    If opt_finalidade_normal Then
        lLinhaDados = lLinhaDados & "1"
    Else
        lLinhaDados = lLinhaDados & "2"
    End If
    Print #1, lLinhaDados
    lRegistro = lRegistro + 1
    'Registro Tipo 11
    'lLinhaDados Complementares do Informante
    lLinhaDados = "11"
    'Logradouro
    'x_endereco = Trim(!Endereco) & " " & Trim(!Bairro)
    
    xString = Space(34)
    If Trim(Empresa.Nome) = "VALPOSTO COMBUSTIVEIS LTDA" Then
        Mid(xString, 1, 34) = "Alameda Ricardo Paranhos"
    Else
        Mid(xString, 1, 34) = Empresa.Endereco
    End If
    lLinhaDados = lLinhaDados & xString
    'Numero
    If Trim(Empresa.Nome) = "VALPOSTO COMBUSTIVEIS LTDA" Then
        lLinhaDados = lLinhaDados & "00504"
    ElseIf UCase(Trim(Empresa.Nome)) = "AUTO POSTO RIO FORMOSO LTDA" Then
        lLinhaDados = lLinhaDados & "00769"
    Else
        lLinhaDados = lLinhaDados & "00000"
    End If
    'Complemento
    '                1234567890123456789012
    lLinhaDados = lLinhaDados & ".                     "
    'Bairro
    xString = Space(15)
    Mid(xString, 1, 15) = Empresa.Bairro
    lLinhaDados = lLinhaDados & xString
    'Cep
    lLinhaDados = lLinhaDados & Empresa.CEP
    'Contato
    xString = Space(28)
    Mid(xString, 1, 28) = Empresa.ResponsavelLegal
    lLinhaDados = lLinhaDados & xString
    'Telefone (tamanho 12 com zero a esquerda)
    xString = "000000000000"
    Mid(xString, 2, 11) = Empresa.Telefone
    lLinhaDados = lLinhaDados & xString
    'If g_empresa = 1 Then
    '    'FONE
    '    lLinhaDados = lLinhaDados & "000622711202"
    Print #1, lLinhaDados
    Exit Sub
ErrorRotina:
    Close #1
    'If Err = 53 Then
    '    MsgBox "Arquivo não encontrado para a data informada.", vbCritical, "Erro no processamento"
    '    Exit Sub
    'Else
        MsgBox "Não foi possível processar o disquete." & Chr(10) & "Erro de número " & Err, vbCritical, "GravaArquivoRegistro10e11"
        Exit Sub
    'End If
End Sub
Private Sub LoopRsEntradaCombustivel(ByVal pTipoRegistro As Integer)
    Dim xData As Date
'    Dim i As Integer
'    Dim i2 As Integer
'    Dim x_inscricao_estadual As String
'
'    Dim xUF As String
'    Dim xValorContabil As Currency
'    Dim xBaseCalculo As Currency
'    Dim xValorOutras As Currency
'    Dim xPorcentagemICMSNormal As Currency
'    Dim xValorICMS As Currency
'    Dim xIcmsNormal As Currency
'    Dim xValorIcmsST As Currency
'    Dim xDescontou As Boolean
    
    On Error GoTo ErrorRotina
    
    xData = CDate(msk_data_i.Text)
    Do Until xData > CDate(msk_data_f.Text)
        
        '********************************
        '**** ENTRADA DE COMBUSTIVEL ****
        '********************************
        'Prepara SQL
        lSQL = ""
        lSQL = lSQL & "SELECT Data, [Numero da Nota], Modelo, Serie, [Codigo do Fornecedor], Sum([Valor da Entrada]) As TOTAL"
        lSQL = lSQL & "  FROM Entrada_Combustivel_LMC"
        lSQL = lSQL & " WHERE Empresa = " & g_empresa
        lSQL = lSQL & "   AND Data = " & preparaData(xData)
        lSQL = lSQL & " GROUP BY Data, [Numero da Nota], Modelo, Serie, [Codigo do Fornecedor]"
        lSQL = lSQL & " ORDER BY Data, [Numero da Nota], Modelo, Serie, [Codigo do Fornecedor]"
    
        'Abre RecordSet
        Set rsTabela = New adodb.Recordset
        Set rsTabela = Conectar.RsConexao(lSQL)
    
        If rsTabela.RecordCount > 0 Then
            Do Until rsTabela.EOF
                If Not Fornecedor.LocalizarCodigo(g_empresa, rsTabela("Codigo do Fornecedor").Value) Then
                    MsgBox "Fornecedor não cadastrado!", vbCritical, "Erro na geração do arquivo!"
                    Exit Sub
                Else
                    lCNPJ = "00000000000000"
                    Mid(lCNPJ, 1, 14) = Fornecedor.CGC
                    lIE = DesmascaraIscricaoEstadual(Fornecedor.InscricaoEstadual)
                    lUF = UCase(Fornecedor.UF)
                End If
                lDataEmissao = Format(rsTabela("Data").Value, "yyyyMMdd")
                lCodModDocFisc = Space(2)
                Mid(lCodModDocFisc, 1, 2) = rsTabela("Modelo").Value
                lSerie = Space(3)
                Mid(lSerie, 1, 3) = rsTabela("Serie").Value
                lNumeroNF = Format(CLng(Trim(rsTabela("Numero da Nota").Value)), "000000")
                lCFOP = "1652"
                lEmitente = "T"
                lValorNF = Mid(Format(rsTabela("TOTAL").Value, "00000000000.00"), 1, 11) & Mid(Format(rsTabela("TOTAL").Value, "00000000000.00"), 13, 2)
                lBaseCalculo = Mid(Format(0, "00000000000.00"), 1, 11) & Mid(Format(0, "00000000000.00"), 13, 2)
                lValorICMS = Mid(Format(0, "00000000000.00"), 1, 11) & Mid(Format(0, "00000000000.00"), 13, 2)
                lIsentas = Mid(Format(0, "00000000000.00"), 1, 11) & Mid(Format(0, "00000000000.00"), 13, 2)
                lOutras = Mid(Format(0, "00000000000.00"), 1, 11) & Mid(Format(0, "00000000000.00"), 13, 2)
                lAliquotaICMS = Mid(Format(0, "00.00"), 1, 2) & Mid(Format(0, "00.00"), 4, 2)
                lNFCancelada = "N"
                If pTipoRegistro = 50 Then
                    Call GravaArquivoRegistro50Ent(rsTabela("Data").Value, lNumeroNF, lSerie)
                ElseIf pTipoRegistro = 53 Then
                    Call GravaArquivoRegistro53Ent(rsTabela("Data").Value, lNumeroNF, lSerie)
                End If
                rsTabela.MoveNext
            Loop
        End If
        rsTabela.Close
        Set rsTabela = Nothing
    
    
        
        '****************************
        '**** ENTRADA DE PRODUTO ****
        '****************************
        'Prepara SQL
        lSQL = ""
        lSQL = lSQL & "SELECT [Data da Entrada], [Numero do Documento], CFOP, Modelo, Serie, [Codigo do Fornecedor], [Total da Nota]"
        lSQL = lSQL & "  FROM EntradaProdutoCabecalho"
        lSQL = lSQL & " WHERE Empresa = " & g_empresa
        lSQL = lSQL & "   AND [Data da Entrada] = " & preparaData(xData)
        lSQL = lSQL & "   AND [Tipo da Entrada] = " & preparaTexto("1")
        lSQL = lSQL & " ORDER BY [Data da Entrada], [Numero do Documento], CFOP, Modelo, Serie, [Codigo do Fornecedor]"
    
        'Abre RecordSet
        Set rsTabela = New adodb.Recordset
        Set rsTabela = Conectar.RsConexao(lSQL)
    
        If rsTabela.RecordCount > 0 Then
            Do Until rsTabela.EOF
                If Not Fornecedor.LocalizarCodigo(g_empresa, rsTabela("Codigo do Fornecedor").Value) Then
                    MsgBox "Fornecedor não cadastrado!", vbCritical, "Erro na geração do arquivo!"
                    Exit Sub
                Else
                    lCNPJ = "00000000000000"
                    Mid(lCNPJ, 1, 14) = Fornecedor.CGC
                    lIE = DesmascaraIscricaoEstadual(Fornecedor.InscricaoEstadual)
                    lUF = UCase(Fornecedor.UF)
                End If
                lDataEmissao = Format(rsTabela("Data da Entrada").Value, "yyyyMMdd")
                lCodModDocFisc = Space(2)
                Mid(lCodModDocFisc, 1, 2) = rsTabela("Modelo").Value
                lSerie = Space(3)
                Mid(lSerie, 1, 3) = rsTabela("Serie").Value
                lNumeroNF = Format(CLng(Trim(rsTabela("Numero do Documento").Value)), "000000")
                lCFOP = rsTabela("CFOP").Value
                lEmitente = "T"
                lValorNF = Mid(Format(rsTabela("Total da Nota").Value, "00000000000.00"), 1, 11) & Mid(Format(rsTabela("Total da Nota").Value, "00000000000.00"), 13, 2)
                lBaseCalculo = Mid(Format(0, "00000000000.00"), 1, 11) & Mid(Format(0, "00000000000.00"), 13, 2)
                lValorICMS = Mid(Format(0, "00000000000.00"), 1, 11) & Mid(Format(0, "00000000000.00"), 13, 2)
                lIsentas = Mid(Format(0, "00000000000.00"), 1, 11) & Mid(Format(0, "00000000000.00"), 13, 2)
                lOutras = Mid(Format(0, "00000000000.00"), 1, 11) & Mid(Format(0, "00000000000.00"), 13, 2)
                lAliquotaICMS = Mid(Format(0, "00.00"), 1, 2) & Mid(Format(0, "00.00"), 4, 2)
                lNFCancelada = "N"
                If pTipoRegistro = 50 Then
                    Call GravaArquivoRegistro50Ent(rsTabela("Data da Entrada").Value, lNumeroNF, lSerie)
                ElseIf pTipoRegistro = 53 Then
                    Call GravaArquivoRegistro53Ent(rsTabela("Data da Entrada").Value, lNumeroNF, lSerie)
                End If
                rsTabela.MoveNext
            Loop
        End If
        rsTabela.Close
        Set rsTabela = Nothing
        
        
        
        
        xData = DateAdd("d", 1, xData)
    Loop
    
    
    
'    '********************************
'    '**** ENTRADA DE COMBUSTIVEL ****
'    '********************************
'    'Prepara SQL
'    lSQL = ""
'    lSQL = lSQL & "SELECT Data, [Numero da Nota], Modelo, Serie, [Codigo do Fornecedor], Sum([Valor da Entrada]) As TOTAL"
'    lSQL = lSQL & "  FROM Entrada_Combustivel_LMC"
'    lSQL = lSQL & " WHERE Empresa = " & g_empresa
'    lSQL = lSQL & "   AND Data >= " & preparaData(msk_data_i.Text)
'    lSQL = lSQL & "   AND Data <= " & preparaData(msk_data_f.Text)
'    lSQL = lSQL & " GROUP BY Data, [Numero da Nota], Modelo, Serie, [Codigo do Fornecedor]"
'    lSQL = lSQL & " ORDER BY Data, [Numero da Nota], Modelo, Serie, [Codigo do Fornecedor]"
'
'    'Abre RecordSet
'    Set rsTabela = New adodb.Recordset
'    Set rsTabela = Conectar.RsConexao(lSQL)
'
'    If rsTabela.RecordCount > 0 Then
'        rsTabela.MoveFirst
'        Do Until rsTabela.EOF
'
'            If Not Fornecedor.LocalizarCodigo(g_empresa, rsTabela("Codigo do Fornecedor").Value) Then
'                MsgBox "Fornecedor não cadastrado!", vbCritical, "Erro na geração do arquivo!"
'                Exit Sub
'            Else
'                lCNPJ = "00000000000000"
'                Mid(lCNPJ, 1, 14) = Fornecedor.CGC
'                lIE = DesmascaraIscricaoEstadual(Fornecedor.InscricaoEstadual)
'                lUF = UCase(Fornecedor.UF)
'            End If
'            lDataEmissao = Format(rsTabela("Data").Value, "yyyyMMdd")
'            lCodModDocFisc = Space(2)
'            Mid(lCodModDocFisc, 1, 2) = rsTabela("Modelo").Value
'            lSerie = Space(3)
'            Mid(lSerie, 1, 3) = rsTabela("Serie").Value
'            lNumeroNF = Format(CLng(Trim(rsTabela("Numero da Nota").Value)), "000000")
'            lCFOP = "1652"
'            lEmitente = "T"
'            lValorNF = Mid(Format(rsTabela("TOTAL").Value, "00000000000.00"), 1, 11) & Mid(Format(rsTabela("TOTAL").Value, "00000000000.00"), 13, 2)
'            lBaseCalculo = Mid(Format(0, "00000000000.00"), 1, 11) & Mid(Format(0, "00000000000.00"), 13, 2)
'            lValorICMS = Mid(Format(0, "00000000000.00"), 1, 11) & Mid(Format(0, "00000000000.00"), 13, 2)
'            lIsentas = Mid(Format(0, "00000000000.00"), 1, 11) & Mid(Format(0, "00000000000.00"), 13, 2)
'            lOutras = Mid(Format(0, "00000000000.00"), 1, 11) & Mid(Format(0, "00000000000.00"), 13, 2)
'            lAliquotaICMS = Mid(Format(0, "00.00"), 1, 2) & Mid(Format(0, "00.00"), 4, 2)
'            lNFCancelada = "N"
'            If pTipoRegistro = 50 Then
'                Call GravaArquivoRegistro50Ent(rsTabela("Data").Value, lNumeroNF, lSerie)
'            ElseIf pTipoRegistro = 53 Then
'                Call GravaArquivoRegistro53Ent(rsTabela("Data").Value, lNumeroNF, lSerie)
'            End If
'            rsTabela.MoveNext
'        Loop
'    End If
'    rsTabela.Close
'    Set rsTabela = Nothing
    
    Exit Sub
    
    
    
    
'    lSQL = "SELECT [Data de Entrada], Numero, CFOP, Desdobramento, Serie, SUM([Valor Total]) AS ValorContabil"
'    lSQL = lSQL & "  FROM Movimento_Nota_Fiscal_Entrada"
'    lSQL = lSQL & " WHERE Empresa = " & g_empresa
'    lSQL = lSQL & "   AND [Data de Entrada] >= " & Chr(35) & Format(CDate(msk_data_i.Text), "mm/dd/yyyy") & Chr(35)
'    lSQL = lSQL & "   AND [Data de Entrada] <= " & Chr(35) & Format(CDate(msk_data_f.Text), "mm/dd/yyyy") & Chr(35)
'    If pTransporte = True Then
'        lSQL = lSQL & "   AND CFOP = " & Chr(39) & "2353" & Chr(39)
'    Else
'        lSQL = lSQL & "   AND CFOP <> " & Chr(39) & "2353" & Chr(39)
'    End If
'    lSQL = lSQL & " GROUP BY [Data de Entrada], Numero, CFOP, Desdobramento, Serie"
'    Set rsEntradaCabecalho = New adodb.Recordset
'    Set rsEntradaCabecalho = Conectar.RsConexao(lSQL)
'
'    If Not rsEntradaCabecalho.BOF Or Not rsEntradaCabecalho.EOF Then
'        Do Until rsEntradaCabecalho.EOF
''            If rsEntradaCabecalho("Numero").Value = 156011 Then
''                MsgBox rsEntradaCabecalho("Numero").Value
''            End If
'            If rsEntradaCabecalho("Serie").Value <> "AC" And rsEntradaCabecalho("Serie").Value <> "CP" Then
'
'                xValorContabil = 0
'                xBaseCalculo = 0
'                xValorOutras = 0
'                xPorcentagemICMSNormal = 0
'                xValorICMS = 0
'                xIcmsNormal = 0
'                xValorIcmsST = 0
'                xDescontou = False
'
'                '+Empresa;+Data de Entrada;+Numero;+Serie
''                tbl_movimento_cabecalho_nota_fiscal_entrada.Seek "=", g_empresa, rsEntradaCabecalho("Data de Entrada").Value, rsEntradaCabecalho("Numero").Value, rsEntradaCabecalho("Serie").Value
''                If Not tbl_movimento_cabecalho_nota_fiscal_entrada.NoMatch Then
''                    With tbl_movimento_nota_fiscal_entrada
''                        tbl_fornecedor.Seek "=", g_empresa, tbl_movimento_cabecalho_nota_fiscal_entrada![Codigo do Fornecedor]
''                        lCNPJ = "00000000000000"
''                        lIE = "ISENTO        "
''                        If Not tbl_fornecedor.NoMatch Then
''                        lCNPJ = tbl_fornecedor!CGC
''                            x_inscricao_estadual = "              "
''                            i2 = 0
''                            For i = 1 To 20
''                                If Mid(tbl_fornecedor![Inscricao Estadual], i, 1) >= "0" And Mid(tbl_fornecedor![Inscricao Estadual], i, 1) <= "9" Then
''                                    i2 = i2 + 1
''                                    Mid(x_inscricao_estadual, i2, 1) = Mid(tbl_fornecedor![Inscricao Estadual], i, 1)
''                                End If
''                            Next
''                            lIE = x_inscricao_estadual
''                        End If
''                        xUF = tbl_movimento_cabecalho_nota_fiscal_entrada!UF
''                        .Seek ">=", g_empresa, rsEntradaCabecalho("Data de Entrada").Value, rsEntradaCabecalho("Numero").Value, rsEntradaCabecalho("Serie").Value, 0
''                        If Not .NoMatch Then
''                            Do Until .EOF
''                                If !Empresa <> g_empresa Or ![Data de Entrada] <> rsEntradaCabecalho("Data de Entrada").Value Or !numero <> rsEntradaCabecalho("Numero").Value Or Trim(!Serie) <> Trim(rsEntradaCabecalho("Serie").Value) Then
''                                    Exit Do
''                                End If
''                                If !CFOP = rsEntradaCabecalho("CFOP").Value Then
''                                    xValorContabil = xValorContabil + ![Valor Total] + ![Valor do IPI]
''                                    If !Desdobramento = "1" Then
''                                        xBaseCalculo = xBaseCalculo + ![Valor Total] + ![Valor do IPI]
''                                        'xValorOutras = ![Aliquota de ICMS]
''                                        xPorcentagemICMSNormal = ![Aliquota de ICMS]
''                                    ElseIf !Desdobramento = "2" Then
''                                        xValorOutras = xValorOutras + ![Valor Total] + ![Valor do IPI]
''                                        xPorcentagemICMSNormal = ![Aliquota de ICMS]
''                                    ElseIf !Desdobramento = "0" Then
''                                        xBaseCalculo = xBaseCalculo + ![Valor Total] + ![Valor do IPI]
''                                        xPorcentagemICMSNormal = 0 '![Aliquota de ICMS]
''                                    End If
''                                    If tbl_movimento_cabecalho_nota_fiscal_entrada![Valor do Desconto] > 0 And xDescontou = False Then
''                                        xDescontou = True
''                                        xValorContabil = xValorContabil - tbl_movimento_cabecalho_nota_fiscal_entrada![Valor do Desconto]
''                                        'If xValorOutras > 0 Then
''                                        '    xValorOutras = xValorOutras - tbl_movimento_cabecalho_nota_fiscal_entrada![Valor do Desconto]
''                                        'End If
''                                        If xBaseCalculo > 0 Then
''                                            xBaseCalculo = xBaseCalculo - tbl_movimento_cabecalho_nota_fiscal_entrada![Valor do Desconto]
''                                        End If
''                                    End If
''                                End If
''                                .MoveNext
''                            Loop
''                        End If
''                        If xUF = "GO" Then
''                            If rsEntradaCabecalho("Data de Entrada").Value < CDate("01/08/2007") Then
''                                xBaseCalculo = tbl_movimento_cabecalho_nota_fiscal_entrada![Base de Calculo do ICMS]
''                            End If
''                            'xPorcentagemICMSNormal = xValorOutras
''                            xValorOutras = xValorContabil - xBaseCalculo
''                        End If
''
''                        If xValorOutras > 0 Then
''                            xValorICMS = Format((xBaseCalculo * xValorOutras / 100), "0000000000.00")
''                        End If
''                        If xPorcentagemICMSNormal > 0 Then
''                            xIcmsNormal = Format((xValorOutras * xPorcentagemICMSNormal / 100), "0000000000.00")
''                            If rsEntradaCabecalho("Data de Entrada").Value >= CDate("01/09/2002") And rsEntradaCabecalho("Data de Entrada").Value <= CDate("30/09/2002") Then
''                                xValorIcmsST = Format((xValorOutras * 1.5) * 12 / 100 - xIcmsNormal, "0000000000.00")
''                            ElseIf rsEntradaCabecalho("Data de Entrada").Value >= CDate("01/10/2002") Then
''                                xValorIcmsST = Format((xValorOutras * 1.4) * 12 / 100 - xIcmsNormal, "0000000000.00")
''                            End If
''                        End If
''
'''                        If rsEntradaCabecalho("CFOP").Value = "2102" Or rsEntradaCabecalho("CFOP").Value = "1102" Then
'''                            lEntradaTributacao = lEntradaTributacao + xValorContabil
'''                            lEntradaTributacaoBC = lEntradaTributacaoBC + xBaseCalculo
'''                            lEntradaTributacaoICMS = lEntradaTributacaoICMS + xValorICMS
'''                            lEntradaTributacaoOutras = lEntradaTributacaoOutras + xValorOutras
'''                            lEntradaTributacaoIcmsNormal = lEntradaTributacaoIcmsNormal
'''                            lEntradaTributacaoIcmsST = lEntradaTributacaoIcmsST + xValorIcmsST
'''                        ElseIf rsEntradaCabecalho("CFOP").Value = "2353" Then
'''                            lEntradaFrete = lEntradaFrete + xValorContabil
'''                            lEntradaFreteBC = lEntradaFreteBC + xBaseCalculo
'''                            lEntradaFreteICMS = lEntradaFreteICMS + xValorICMS
'''                            lEntradaFreteOutras = lEntradaFreteOutras + xValorOutras
'''                            lEntradaFreteIcmsNormal = lEntradaFreteIcmsNormal + xIcmsNormal
'''                            lEntradaFreteIcmsST = lEntradaFreteIcmsST + xValorIcmsST
'''                        ElseIf rsEntradaCabecalho("CFOP").Value = "2403" Then
'''                            lEntradaSubstituicao = lEntradaSubstituicao + xValorContabil
'''                            lEntradaSubstituicaoBC = lEntradaSubstituicaoBC + xBaseCalculo
'''                            lEntradaSubstituicaoICMS = lEntradaSubstituicaoICMS + xValorICMS
'''                            lEntradaSubstituicaoOutras = lEntradaSubstituicaoOutras + xValorOutras
'''                            lEntradaSubstituicaoIcmsNormal = lEntradaSubstituicaoIcmsNormal + xIcmsNormal
'''                            lEntradaSubstituicaoIcmsST = lEntradaSubstituicaoIcmsST + xValorIcmsST
'''                        End If
''
''                    End With
''                    lLinhaDados = "50"
''                    lDataEmissao = Mid(rsEntradaCabecalho("Data de Entrada").Value, 7, 4) & Mid(rsEntradaCabecalho("Data de Entrada").Value, 4, 2) & Mid(rsEntradaCabecalho("Data de Entrada").Value, 1, 2)
''                    lUF = tbl_fornecedor!UF
''                    lCodModDocFisc = "01"
''                    lSerie = Space(3)
''                    Mid(lSerie, 1, 2) = rsEntradaCabecalho("Serie").Value
''                    lSubSerie = "  "
''                    lNumeroNF = Mid(Format(rsEntradaCabecalho("Numero").Value, "000000"), 1, 6)
''                    lCFOP = rsEntradaCabecalho("CFOP").Value
''                    '''If lCFOP = "2353" Then
''                    '''    lCFOP = "2102"
''                    '''End If
''                    lEmitente = "P"
''                    lValorNF = Mid(Format(xValorContabil, "00000000000.00"), 1, 11) & Mid(Format(xValorContabil, "00000000000.00"), 13, 2)
''                    If rsEntradaCabecalho("Data de Entrada").Value >= CDate("01/08/2007") Then
''                        xBaseCalculo = 0
''                        xValorOutras = xValorContabil
''                        xPorcentagemICMSNormal = 0
''                    End If
''                    lBaseCalculo = Mid(Format(xBaseCalculo, "00000000000.00"), 1, 11) & Mid(Format(xBaseCalculo, "00000000000.00"), 13, 2)
''                    lValorICMS = Mid(Format(xValorICMS, "00000000000.00"), 1, 11) & Mid(Format(xValorICMS, "00000000000.00"), 13, 2)
''                    lIsentas = Mid(Format(0, "00000000000.00"), 1, 11) & Mid(Format(0, "00000000000.00"), 13, 2)
''                    lOutras = Mid(Format(xValorOutras, "00000000000.00"), 1, 11) & Mid(Format(xValorOutras, "00000000000.00"), 13, 2)
''                    '''lAliquotaICMS = Mid(Format(tbl_movimento_cabecalho_nota_fiscal_entrada![Porcentagem do ICMS], "00.00"), 1, 2) & Mid(Format(tbl_movimento_cabecalho_nota_fiscal_entrada![Porcentagem do ICMS], "00.00"), 4, 2)
''                    If rsEntradaCabecalho("Data de Entrada").Value < CDate("01/08/2007") Then
''                        If xPorcentagemICMSNormal = 0 Then
''                            xPorcentagemICMSNormal = xValorOutras
''                        End If
''                    End If
''                    lAliquotaICMS = Mid(Format(xPorcentagemICMSNormal, "00.00"), 1, 2) & Mid(Format(xPorcentagemICMSNormal, "00.00"), 4, 2)
''                    lNFCancelada = "N"
''                    'Monta String "lLinhaDados"
''                    If pTransporte = False Then
''                        Call GravaArquivoRegistro53Ent(rsEntradaCabecalho("Data de Entrada").Value, rsEntradaCabecalho("Numero").Value, rsEntradaCabecalho("Serie").Value)
''                    Else
''                        Call GravaLinhaDadosDisquete70Ent(rsEntradaCabecalho("Data de Entrada").Value, rsEntradaCabecalho("Numero").Value, rsEntradaCabecalho("Serie").Value)
''                    End If
''                Else
''                    MsgBox "Cabeçalho de NF de Entrada não encontrada!", vbInformation, "Erro de Integridade!"
''                End If
'            End If
'            rsEntradaCabecalho.MoveNext
'        Loop
'    End If
'    rsEntradaCabecalho.Close
'    Set rsEntradaCabecalho = Nothing
    
    
    
'    With tbl_movimento_cabecalho_nota_fiscal_entrada
'        .Seek ">=", g_empresa, CDate(msk_data_i.Text), 0, "  "
'        If Not .NoMatch Then
'            Do Until .EOF
'                If !Empresa <> g_empresa Or ![Data de Entrada] > CDate(msk_data_f.Text) Then
'                    Exit Do
'                End If
'                If !Serie <> "AC" And !Serie <> "CP" And ((pTransporte = False And ![Codificacao Fiscal] <> "2353") Or (pTransporte = True And ![Codificacao Fiscal] = "2353")) Then
'                    'Registro Tipo 50 (Entradas)
'                    lLinhaDados = "50"
'                    If ![Codigo do Fornecedor] > 0 Then
'                        tbl_fornecedor.Seek "=", g_empresa, ![Codigo do Fornecedor]
'                        If tbl_fornecedor.NoMatch Then
'                            MsgBox "Fornecedor não cadastrado!", vbCritical, "Erro na geração do disquete!"
'                            Exit Sub
'                        End If
'                        lCNPJ = tbl_fornecedor!CGC
'                        x_inscricao_estadual = "              "
'                        i2 = 0
'                        For i = 1 To 20
'                            If Mid(tbl_fornecedor![Inscricao Estadual], i, 1) >= "0" And Mid(tbl_fornecedor![Inscricao Estadual], i, 1) <= "9" Then
'                                i2 = i2 + 1
'                                Mid(x_inscricao_estadual, i2, 1) = Mid(tbl_fornecedor![Inscricao Estadual], i, 1)
'                            End If
'                        Next
'                        lIE = x_inscricao_estadual
'                    Else
'                        lCNPJ = "00000000000000"
'                        lIE = "ISENTO        "
'                    End If
'                    lDataEmissao = Mid(![Data de Entrada], 7, 4) & Mid(![Data de Entrada], 4, 2) & Mid(![Data de Entrada], 1, 2)
'                    lUF = tbl_fornecedor!UF
'                    lCodModDocFisc = "01"
'                    lSerie = Space(3)
'                    Mid(lSerie, 1, 2) = !Serie
'                    lSubSerie = "  "
'                    lNumeroNF = Mid(Format(!numero, "000000"), 1, 6)
'                    lCFOP = ![Codificacao Fiscal]
'                    '''If lCFOP = "2353" Then
'                    '''    lCFOP = "2102"
'                    '''End If
'                    lEmitente = "P"
'                    'lValorNF = Mid(Format(![Total da Nota] + ![Valor do Desconto], "00000000000.00"), 1, 11) & Mid(Format(![Total da Nota] + ![Valor do Desconto], "00000000000.00"), 13, 2)
'                    lValorNF = Mid(Format(![Total da Nota], "00000000000.00"), 1, 11) & Mid(Format(![Total da Nota], "00000000000.00"), 13, 2)
'                    lBaseCalculo = Mid(Format(![Base de Calculo do ICMS], "00000000000.00"), 1, 11) & Mid(Format(![Base de Calculo do ICMS], "00000000000.00"), 13, 2)
'                    'If ![Porcentagem do ICMS] > 0 Then
'                    '    lBaseCalculo = Mid(Format(![Total da Nota], "00000000000.00"), 1, 11) & Mid(Format(![Total da Nota], "00000000000.00"), 13, 2)
'                    'Else
'                    '    lBaseCalculo = Mid(Format(0, "00000000000.00"), 1, 11) & Mid(Format(0, "00000000000.00"), 13, 2)
'                    'End If
'                    lValorICMS = Mid(Format(![Valor do ICMS], "00000000000.00"), 1, 11) & Mid(Format(![Valor do ICMS], "00000000000.00"), 13, 2)
'                    lIsentas = Mid(Format(0, "00000000000.00"), 1, 11) & Mid(Format(0, "00000000000.00"), 13, 2)
'                    If ![Porcentagem do ICMS] > 0 Then
'                        lOutras = Mid(Format(![Valor de Outros], "00000000000.00"), 1, 11) & Mid(Format(![Valor de Outros], "00000000000.00"), 13, 2)
'                    Else
'                        lOutras = Mid(Format(![Total da Nota], "00000000000.00"), 1, 11) & Mid(Format(![Total da Nota], "00000000000.00"), 13, 2)
'                    End If
'                    lAliquotaICMS = Mid(Format(![Porcentagem do ICMS], "00.00"), 1, 2) & Mid(Format(![Porcentagem do ICMS], "00.00"), 4, 2)
'                    lNFCancelada = "N"
'                    'Monta String "lLinhaDados"
'                    If pTransporte = False Then
'                        Call GravaArquivoRegistro53Ent(![Data de Entrada], !numero, !Serie)
'                    Else
'                        Call GravaLinhaDadosDisquete70Ent(![Data de Entrada], !numero, !Serie)
'                    End If
'                End If
'                .MoveNext
'            Loop
'        End If
'    End With
    Exit Sub
ErrorRotina:
    Close #1
    MsgBox "Não foi possível processar o arquivo." & Chr(10) & "Erro de número " & Err, vbCritical, "LoopRsEntradaCombustivel"
    Exit Sub
End Sub
Private Sub LoopRsMapaResumo()
    
    On Error GoTo ErrorRotina
    
    '********************************
    '**** Mapa Resumo            ****
    '********************************
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Data, Numero, [ECF Numero], [Contagem de Operacao Inicial], "
    lSQL = lSQL & "       [Contagem de Operacao Final], [Totalizador Geral Final], "
    lSQL = lSQL & "       [Totalizador Geral Inicial], Isentas, [Nao Incidencia], "
    lSQL = lSQL & "       [Substituicao Tributaria], [ICMS 17], [Cancelamento de Item], "
    lSQL = lSQL & "       [Contador de Reducoes Z], [ICMS 12], [Contagem de Reinicio de Operacao], "
    lSQL = lSQL & "       Desconto, Acrescimo"
    lSQL = lSQL & "  FROM Mapa_Resumo"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND Data <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & " ORDER BY Data, [ECF Numero]"

    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    
    If rsTabela.RecordCount > 0 Then
        rsTabela.MoveFirst
        Do Until rsTabela.EOF
            Call GravaArquivoRegistro60M
            If rsTabela("Cancelamento de Item").Value > 0 Then
                Call GravaArquivoRegistro60A("CANC", 0, rsTabela("Cancelamento de Item").Value)
            End If
            If rsTabela("Desconto").Value > 0 Then
                Call GravaArquivoRegistro60A("DESC", 0, rsTabela("Desconto").Value)
            End If
            If rsTabela("Substituicao Tributaria").Value > 0 Then
                Call GravaArquivoRegistro60A("F   ", 0, rsTabela("Substituicao Tributaria").Value)
            End If
            If rsTabela("Isentas").Value > 0 Then
                Call GravaArquivoRegistro60A("I   ", 0, rsTabela("Isentas").Value)
            End If
            If rsTabela("Nao Incidencia").Value > 0 Then
                Call GravaArquivoRegistro60A("N   ", 0, rsTabela("Nao Incidencia").Value)
            End If
            If rsTabela("ICMS 17").Value > 0 Then
                Call GravaArquivoRegistro60A("Tributada", 17, rsTabela("ICMS 17").Value)
            End If
            rsTabela.MoveNext
        Loop
    End If
    rsTabela.Close
    Set rsTabela = Nothing
    Exit Sub

ErrorRotina:
    Close #1
    MsgBox "Não foi possível processar o arquivo." & Chr(10) & "Erro de número " & Err, vbCritical, "LoopRsMapaResumo"
    Exit Sub
End Sub
Private Sub GravaArquivoRegistro50Ent(ByVal pData As Date, ByVal pNumero As Long, ByVal pSerie As String)
    
    On Error GoTo ErrorRotina
    
    lLinhaDados = "50"
    If lCFOP = "2353" Then
        lLinhaDados = "70"
    End If
    lLinhaDados = lLinhaDados & lCNPJ
    lLinhaDados = lLinhaDados & lIE
    lLinhaDados = lLinhaDados & lDataEmissao
    lLinhaDados = lLinhaDados & lUF
    lLinhaDados = lLinhaDados & lCodModDocFisc
    If lSerie = "U  " Then
        lLinhaDados = lLinhaDados & "1  "
    Else
        lLinhaDados = lLinhaDados & lSerie
    End If
    lLinhaDados = lLinhaDados & lNumeroNF
    lLinhaDados = lLinhaDados & lCFOP
    lLinhaDados = lLinhaDados & lEmitente
    lLinhaDados = lLinhaDados & lValorNF
    lLinhaDados = lLinhaDados & lBaseCalculo
    lLinhaDados = lLinhaDados & lValorICMS
    lLinhaDados = lLinhaDados & lIsentas
    lLinhaDados = lLinhaDados & lOutras
    lLinhaDados = lLinhaDados & lAliquotaICMS
    lLinhaDados = lLinhaDados & lNFCancelada
    'Grava String "lLinhaDados" no Arquivo Texto
    Print #1, lLinhaDados
    lRegistro = lRegistro + 1
    If lCFOP = "2353" Then
        lRegistro70 = lRegistro70 + 1
    Else
        lRegistro50 = lRegistro50 + 1
    End If
    'Next
    Exit Sub

ErrorRotina:
    Close #1
    MsgBox "Não foi possível gravar registro tipo 50." & Chr(10) & "Erro de número " & Err, vbCritical, "GravaArquivoRegistro50Ent"
    Exit Sub
End Sub
Private Sub GravaArquivoRegistro53Ent(ByVal pData As Date, ByVal pNumero As Long, ByVal pSerie As String)
    'Dim xValorIPI As Currency
    'Dim xValor As Currency
    
    On Error GoTo ErrorRotina
    
    'xValorIPI = 0
    'With tbl_movimento_nota_fiscal_entrada
    '    .Seek ">=", g_empresa, x_data, x_numero, x_serie, 0
    '    If Not .NoMatch Then
    '        Do Until .EOF
    '            If !Empresa <> g_empresa Or ![Data de Entrada] <> x_data Or !numero <> x_numero Or !Serie <> x_serie Then
    '                Exit Do
    '            End If
    '            xValorIPI = xValorIPI + ![Valor do IPI]
    '            .MoveNext
    '        Loop
    '    End If
    'End With
    'xValor = Mid(lValorNF, 1, 11) & "," & Mid(lValorNF, 12, 2)
    'xValor = xValor - xValorIPI
    'lBaseCalculo = Mid(Format(xValor, "00000000000.00"), 1, 11) & Mid(Format(xValor, "00000000000.00"), 13, 2)

    lLinhaDados = "53"
    If lCFOP = "2353" Then
        lLinhaDados = "70"
    End If
    lLinhaDados = lLinhaDados & lCNPJ
    lLinhaDados = lLinhaDados & lIE
    lLinhaDados = lLinhaDados & lDataEmissao
    lLinhaDados = lLinhaDados & lUF
    lLinhaDados = lLinhaDados & lCodModDocFisc
    If lSerie = "U  " Then
        lLinhaDados = lLinhaDados & "1  "
    Else
        lLinhaDados = lLinhaDados & lSerie
    End If
    lLinhaDados = lLinhaDados & lNumeroNF
    lLinhaDados = lLinhaDados & lCFOP
    lLinhaDados = lLinhaDados & lEmitente
    'lLinhaDados = lLinhaDados & lValorNF
    lLinhaDados = lLinhaDados & lBaseCalculo
    lLinhaDados = lLinhaDados & lValorICMS
    lLinhaDados = lLinhaDados & lIsentas 'Despesas Acessórias
    'lLinhaDados = lLinhaDados & lOutras
    'lLinhaDados = lLinhaDados & lAliquotaICMS
    lLinhaDados = lLinhaDados & lNFCancelada
    lLinhaDados = lLinhaDados & "1" 'Código de Antecipação
    lLinhaDados = lLinhaDados & Space(29) 'Branco (29)
    'Grava String "lLinhaDados" no Arquivo Texto
    Print #1, lLinhaDados
    lRegistro = lRegistro + 1
    If lCFOP = "2353" Then
        lRegistro70 = lRegistro70 + 1
    Else
        lRegistro53 = lRegistro53 + 1
    End If
    'Next
    Exit Sub

ErrorRotina:
    Close #1
    MsgBox "Não foi possível gravar registro tipo 53." & Chr(10) & "Erro de número " & Err, vbCritical, "GravaArquivoRegistro53Ent"
    Exit Sub
End Sub
Private Sub GravaLinhaDadosDisquete70Ent(x_data As Date, x_numero As Long, x_serie As String)
    lLinhaDados = "70"
    lLinhaDados = lLinhaDados & lCNPJ
    lLinhaDados = lLinhaDados & lIE
    lLinhaDados = lLinhaDados & lDataEmissao
    lLinhaDados = lLinhaDados & lUF
    'Transporte Rodoviário
    lLinhaDados = lLinhaDados & "08"
    lLinhaDados = lLinhaDados & lSerie
    lLinhaDados = lLinhaDados & lNumeroNF
    lLinhaDados = lLinhaDados & lCFOP
    'lLinhaDados = lLinhaDados & lEmitente
    lLinhaDados = lLinhaDados & lValorNF
    lLinhaDados = lLinhaDados & "0" & lBaseCalculo
    lLinhaDados = lLinhaDados & "0" & lValorICMS
    lLinhaDados = lLinhaDados & "0" & lIsentas
    lLinhaDados = lLinhaDados & "0" & lOutras
    'CIF FOB
    lLinhaDados = lLinhaDados & "1"
    'lLinhaDados = lLinhaDados & lAliquotaICMS
    lLinhaDados = lLinhaDados & lNFCancelada
    'Grava String "lLinhaDados" no Arquivo Texto
    Print #1, lLinhaDados
    lRegistro = lRegistro + 1
    lRegistro70 = lRegistro70 + 1
    'Next
    Exit Sub
ErrorGravaLinhaDadosDisquete50Sai:
    Close #1
    MsgBox "Não foi possível gravar registro tipo 50." & Chr(10) & "Erro de número " & Err, vbCritical, "GravaLinhaDadosDisquete50Sai"
    Exit Sub
End Sub
Private Sub GravaLinhaDadosDisquete50Sai(x_data As Date, x_impressora As Integer, x_numero As Long, x_serie As String, xQuantidade As Integer)
    Dim i As Integer
    Dim xReducaoBC As Boolean
    Dim xPercentualAliquota(1 To 6) As String
    Dim xValorNF(1 To 6) As Currency
    Dim xBaseCalculo(1 To 6) As Currency
    Dim xOutras(1 To 6) As Currency
    
    Dim xValorReducao As Currency
    Dim xValorICMS As Currency
    Dim i2 As Integer
    On Error GoTo ErrorGravaLinhaDadosDisquete50Sai
'    With tbl_movimento_cabecalho_nota_fiscal_saida
'        xReducaoBC = False
'        If ![Codigo do Cliente] > 0 Then
'            tbl_cliente.Seek "=", ![Codigo do Cliente]
'            If Not tbl_cliente.NoMatch Then
'                If tbl_cliente!UF = g_uf_empresa Then
'                    If Val(tbl_cliente![Inscricao Estadual]) > 0 Then
'                        xReducaoBC = True
'                    End If
'                End If
'            End If
'        End If
'    End With
'    If xQuantidade > 1 Then
'        With tbl_movimento_nota_fiscal_saida
'            .Seek ">=", g_empresa, x_data, x_impressora, x_numero, x_serie, 0
'            If Not .NoMatch Then
'                Do Until .EOF
'                    If !Empresa <> g_empresa Or !Data <> x_data Or !Impressora <> x_impressora Or !numero <> x_numero Or !Serie <> x_serie Then
'                        Exit Do
'                    End If
'                    xValorReducao = 0
'                    If xReducaoBC And ![Aliquota de ICMS] > 0 And ![Codigo da Aliquota] = 4 Then
'                        xValorReducao = Format(![Valor Total] * 16.67 / 100, "0000000000.00")
'                    End If
'                    For i2 = 1 To 6
'                        If xPercentualAliquota(i2) = "" Then
'                            xPercentualAliquota(i2) = Format(![Aliquota de ICMS], "00.00")
'                            xValorNF(i2) = ![Valor Total]
'                            If ![Aliquota de ICMS] > 0 Then
'                                xBaseCalculo(i2) = ![Valor Total] - xValorReducao
'                                xOutras(i2) = xValorReducao
'                            Else
'                                xOutras(i2) = ![Valor Total]
'                            End If
'                            Exit For
'                        ElseIf xPercentualAliquota(i2) = Format(![Aliquota de ICMS], "00.00") Then
'                            xValorNF(i2) = xValorNF(i2) + ![Valor Total]
'                            If ![Aliquota de ICMS] > 0 Then
'                                xBaseCalculo(i2) = xBaseCalculo(i2) + ![Valor Total] - xValorReducao
'                                xOutras(i2) = xOutras(i2) + xValorReducao
'                            Else
'                                xOutras(i2) = xOutras(i2) + ![Valor Total]
'                            End If
'                            Exit For
'                        End If
'                    Next
'                    .MoveNext
'                Loop
'            End If
'        End With
'    End If
    For i = 1 To xQuantidade
        If xQuantidade > 1 Then
            lValorNF = Mid(Format(xValorNF(i), "00000000000.00"), 1, 11) & Mid(Format(xValorNF(i), "00000000000.00"), 13, 2)
            lBaseCalculo = Mid(Format(xBaseCalculo(i), "00000000000.00"), 1, 11) & Mid(Format(xBaseCalculo(i), "00000000000.00"), 13, 2)
            lOutras = Mid(Format(xOutras(i), "00000000000.00"), 1, 11) & Mid(Format(xOutras(i), "00000000000.00"), 13, 2)
            lAliquotaICMS = Mid(xPercentualAliquota(i), 1, 2) & Mid(xPercentualAliquota(i), 4, 2)
            xValorICMS = Format(xBaseCalculo(i) * fValidaValor(CStr(xPercentualAliquota(i))) / 100, "00000000000.00")
            lValorICMS = Mid(Format(xValorICMS, "00000000000.00"), 1, 11) & Mid(Format(xValorICMS, "00000000000.00"), 13, 2)
        End If
        
        lLinhaDados = "50"
        lLinhaDados = lLinhaDados & lCNPJ
        lLinhaDados = lLinhaDados & lIE
        lLinhaDados = lLinhaDados & lDataEmissao
        lLinhaDados = lLinhaDados & lUF
        lLinhaDados = lLinhaDados & lCodModDocFisc
        lLinhaDados = lLinhaDados & lSerie
        lLinhaDados = lLinhaDados & lNumeroNF
        lLinhaDados = lLinhaDados & lCFOP
        lLinhaDados = lLinhaDados & lEmitente
        lLinhaDados = lLinhaDados & lValorNF
        lLinhaDados = lLinhaDados & lBaseCalculo
        lLinhaDados = lLinhaDados & lValorICMS
        lLinhaDados = lLinhaDados & lIsentas
        lLinhaDados = lLinhaDados & lOutras
        lLinhaDados = lLinhaDados & lAliquotaICMS
        lLinhaDados = lLinhaDados & lNFCancelada
        'Grava String "lLinhaDados" no Arquivo Texto
        Print #1, lLinhaDados
        lRegistro = lRegistro + 1
        lRegistro50 = lRegistro50 + 1
    Next
    Exit Sub
ErrorGravaLinhaDadosDisquete50Sai:
    Close #1
    MsgBox "Não foi possível gravar registro tipo 50." & Chr(10) & "Erro de número " & Err, vbCritical, "GravaLinhaDadosDisquete50Sai"
    Exit Sub
End Sub
Private Sub GravaArquivoRegistro50Saida()
    Dim i As Integer
    Dim i2 As Integer
    Dim x_inscricao_estadual As String * 14
    Dim x_aliquota_icms As Currency
    Dim x_valor_icms As Currency
    Dim xQuantidadeAliquota As Integer
    
    On Error GoTo ErrorRotina
    
'    With tbl_movimento_cabecalho_nota_fiscal_saida
'        .Seek ">=", g_empresa, CDate(msk_data_i.Text), 0, 0, "  "
'        If Not .NoMatch Then
'            Do Until .EOF
'                xQuantidadeAliquota = 1
'                If !Empresa <> g_empresa Or !Data > CDate(msk_data_f) Then
'                    Exit Do
'                End If
'                'If !numero = 1118 Then
'                '    MsgBox !numero
'                'End If
'                If !Serie <> "AC" And !Serie <> "CP" And ![Situacao da Venda] <= 2 Then
'                    xQuantidadeAliquota = QuantidadeAliquotaNFSaida(!Data, !Impressora, !numero, !Serie)
'                    'Registro Tipo 50 (Saidas)
'                    If !Cancelada Then
'                        lCNPJ = "00000000000000"
'                        lIE = "              "
'                        lDataEmissao = Mid(!Data, 7, 4) & Mid(!Data, 4, 2) & Mid(!Data, 1, 2)
'                        lUF = "  "
'                        lCodModDocFisc = "01"
'                        lSerie = !Serie & " "
'                        lSubSerie = "  "
'                        lNumeroNF = Format(!numero, "000000")
'                        lCFOP = "0000"
'                        lEmitente = "P"
'                        lValorNF = "0000000000000"
'                        lBaseCalculo = "0000000000000"
'                        lValorICMS = "0000000000000"
'                        lIsentas = "0000000000000"
'                        lOutras = "0000000000000"
'                        lAliquotaICMS = "0000"
'                        lNFCancelada = "S"
'                    Else
'                        lDataEmissao = Mid(!Data, 7, 4) & Mid(!Data, 4, 2) & Mid(!Data, 1, 2)
'                        'Abaixo é um teste para parar uma nota fiscal especifica para debugar
'                        'If !numero = 5069 Or !numero = 5056 Then
'                        '    MsgBox "teste"
'                        'End If
'
'                        If ![Codigo do Cliente] > 0 Then
'                            tbl_cliente.Seek "=", ![Codigo do Cliente]
'                            If Not tbl_cliente.NoMatch Then
'                                lUF = tbl_cliente!UF
'                                If Len(Trim(tbl_cliente!CGC)) > 0 Then
'                                    lCNPJ = tbl_cliente!CGC
'                                ElseIf Len(Trim(tbl_cliente!CPF)) > 0 Then
'                                    lCNPJ = "000" & tbl_cliente!CPF
'                                Else
'                                    lCNPJ = "00000000000000"
'                                End If
'                                If Len(Trim(tbl_cliente![Inscricao Estadual])) > 0 Then
'                                    x_inscricao_estadual = "              "
'                                    i2 = 0
'                                    For i = 1 To 20
'                                        If Mid(tbl_cliente![Inscricao Estadual], i, 1) >= "0" And Mid(tbl_cliente![Inscricao Estadual], i, 1) <= "9" Then
'                                            i2 = i2 + 1
'                                            Mid(x_inscricao_estadual, i2, 1) = Mid(tbl_cliente![Inscricao Estadual], i, 1)
'                                        End If
'                                    Next
'                                    If Len(Trim(x_inscricao_estadual)) > 0 Then
'                                        lIE = x_inscricao_estadual
'                                    Else
'                                        lIE = "ISENTO        "
'                                    End If
'                                Else
'                                    lIE = "ISENTO        "
'                                End If
'                            Else
'                                MsgBox "Cliente não cadastrado!" & Chr(10) & "Código: " & ![Codigo do Cliente] & Chr(10) & "Nota de Saida número: " & !numero, vbCritical, "Erro na geração do disquete!"
'                                lUF = !Estado
'                            End If
'                        Else
'                            lCNPJ = "00000000000000"
'                            lIE = "ISENTO        "
'                            lUF = !Estado
'                        End If
'                        lCodModDocFisc = "01"
'                        lSerie = !Serie & " "
'                        lSubSerie = "  "
'                        lNumeroNF = Format(!numero, "000000")
'                        lCFOP = ![Codificacao Fiscal]
'                        lEmitente = "P"
'                        lValorNF = Mid(Format(![Total da Nota], "00000000000.00"), 1, 11) & Mid(Format(![Total da Nota], "00000000000.00"), 13, 2)
'                        If !Data >= CDate("01/08/2007") Then
'                            x_aliquota_icms = 0
'                        Else
'                            x_aliquota_icms = AliquotaICMS(!Data, !Impressora, !numero, !Serie)
'                        End If
'                        If x_aliquota_icms > 0 Then
'                            lBaseCalculo = Mid(Format(![Base de Calculo do ICMS], "00000000000.00"), 1, 11) & Mid(Format(![Base de Calculo do ICMS], "00000000000.00"), 13, 2)
'                        Else
'                            lBaseCalculo = Mid(Format(0, "00000000000.00"), 1, 11) & Mid(Format(0, "00000000000.00"), 13, 2)
'                        End If
'                        x_valor_icms = ValorICMS(!Data, !Impressora, !numero, !Serie)
'                        lValorICMS = Mid(Format(x_valor_icms, "00000000000.00"), 1, 11) & Mid(Format(x_valor_icms, "00000000000.00"), 13, 2)
'                        lIsentas = Mid(Format(0, "00000000000.00"), 1, 11) & Mid(Format(0, "00000000000.00"), 13, 2)
'                        If !Data >= CDate("01/08/2007") Then
'                            lOutras = Mid(Format(![Total da Nota], "00000000000.00"), 1, 11) & Mid(Format(![Total da Nota], "00000000000.00"), 13, 2)
'                        Else
'                            If x_aliquota_icms > 0 Then
'                                lOutras = Mid(Format(0, "00000000000.00"), 1, 11) & Mid(Format(0, "00000000000.00"), 13, 2)
'                            Else
'                                lOutras = Mid(Format(![Total da Nota], "00000000000.00"), 1, 11) & Mid(Format(![Total da Nota], "00000000000.00"), 13, 2)
'                            End If
'                        End If
'                        lAliquotaICMS = Mid(Format(x_aliquota_icms, "00.00"), 1, 2) & Mid(Format(x_aliquota_icms, "00.00"), 4, 2)
'                        lNFCancelada = "N"
'                    End If
'                    'Monta String "lLinhaDados"
'                    Call GravaLinhaDadosDisquete50Sai(!Data, !Impressora, !numero, !Serie, xQuantidadeAliquota)
'                End If
'                .MoveNext
'            Loop
'        End If
'    End With
    Exit Sub
ErrorRotina:
    Close #1
    'If Err = 53 Then
    '    MsgBox "Arquivo não encontrado para a data informada.", vbCritical, "Erro no processamento"
    '    Exit Sub
    'Else
        MsgBox "Não foi possível processar o disquete." & Chr(10) & "Erro de número " & Err, vbCritical, "GravaArquivoRegistro50Saida"
        Exit Sub
    'End If
End Sub
Private Sub GravaArquivoRegistro54()
    
    On Error GoTo ErrorRotina
    
    With rsItemEntradaSaida
        .Sort = "CNPJ, Serie, Numero, Item"
        .MoveFirst
        Do Until .EOF
            'Registro Tipo 54 (Item de NF Entrada/Saida)
            lLinhaDados = "54"
            lLinhaDados = lLinhaDados & !CNPJ
            lLinhaDados = lLinhaDados & !Modelo
            If !Serie = "U  " Then
                lLinhaDados = lLinhaDados & "1  "
            Else
                lLinhaDados = lLinhaDados & !Serie
            End If
            lLinhaDados = lLinhaDados & !numero
            lLinhaDados = lLinhaDados & !CFOP
            lLinhaDados = lLinhaDados & !CST
            lLinhaDados = lLinhaDados & !Item
            lLinhaDados = lLinhaDados & !CodigoProduto
            lLinhaDados = lLinhaDados & !Quantidade
            lLinhaDados = lLinhaDados & !ValorProduto
            lLinhaDados = lLinhaDados & !ValorDesconto
            lLinhaDados = lLinhaDados & !BaseICMS
            If !BaseICMS = 0 Then
                lLinhaDados = lLinhaDados & !BaseICMS
            Else
                lLinhaDados = lLinhaDados & "000000000000"
            End If
            'lLinhaDados = lLinhaDados & !BaseICMSST
            lLinhaDados = lLinhaDados & !ValorIPI
            If CDate(msk_data_i.Text) >= CDate("01/08/2007") Then
                lLinhaDados = lLinhaDados & "0000"
            Else
                lLinhaDados = lLinhaDados & !Aliquota
            End If
            Print #1, lLinhaDados
            lRegistro = lRegistro + 1
            lRegistro54 = lRegistro54 + 1
            .MoveNext
        Loop
    End With
    Exit Sub

ErrorRotina:
    Close #1
    'If Err = 53 Then
    '    MsgBox "Arquivo não encontrado para a data informada.", vbCritical, "Erro no processamento"
    '    Exit Sub
    'Else
        MsgBox "Não foi possível processar o arquivo." & Chr(10) & "Erro de número " & Err, vbCritical, "GravaArquivoRegistro54"
        Exit Sub
    'End If
End Sub
Private Sub GravaArquivoRegistro60A(ByVal pTipoAliquota As String, ByVal pAliquota As Currency, ByVal pValor As Currency)
    
    On Error GoTo ErrorRotina
    
    lLinhaDados = "60"
    lLinhaDados = lLinhaDados & "A"
    lLinhaDados = lLinhaDados & Format(rsTabela("Data").Value, "yyyyMMdd")
    
    lLinhaDados = lLinhaDados & lNumeroSerieECF(Val(rsTabela("ECF Numero").Value))
    
    'lLinhaDados = lLinhaDados & Format(rsTabela("ECF Numero").Value, "000")
    If pTipoAliquota = "Tributada" Then
        lLinhaDados = lLinhaDados & Mid(Format(pAliquota, "00.00"), 1, 2) & Mid(Format(pAliquota, "00.00"), 4, 2)
    Else
        lLinhaDados = lLinhaDados & pTipoAliquota
    End If
    lLinhaDados = lLinhaDados & Mid(Format(pValor, "0000000000.00"), 1, 10) & Mid(Format(pValor, "0000000000.00"), 12, 2)
    lLinhaDados = lLinhaDados & Space(79) '79 brancos
    
    'Grava String "lLinhaDados" no Arquivo Texto
    Print #1, lLinhaDados
    lRegistro = lRegistro + 1
    lRegistro60A = lRegistro60A + 1
    Exit Sub

ErrorRotina:
    Close #1
    MsgBox "Não foi possível gravar registro tipo 60A." & Chr(10) & "Erro de número " & Err, vbCritical, "GravaArquivoRegistro60A"
    Exit Sub
End Sub
Private Sub GravaArquivoRegistro60M()
    
    On Error GoTo ErrorRotina
    
    lLinhaDados = "60"
    lLinhaDados = lLinhaDados & "M"
    lLinhaDados = lLinhaDados & Format(rsTabela("Data").Value, "yyyyMMdd")
    lLinhaDados = lLinhaDados & lNumeroSerieECF(Val(rsTabela("ECF Numero").Value))
    lLinhaDados = lLinhaDados & Format(rsTabela("ECF Numero").Value, "000")
    
    ' Modelo do Documento Fiscal
    ' "2D", quando se tratar de Cupom Fiscal (emitido por ECF)
    lLinhaDados = lLinhaDados & "2D"
    lLinhaDados = lLinhaDados & Format(rsTabela("Contagem de Operacao Inicial").Value, "000000")
    lLinhaDados = lLinhaDados & Format(rsTabela("Contagem de Operacao Final").Value, "000000")
    lLinhaDados = lLinhaDados & Format(rsTabela("Contador de Reducoes Z").Value, "000000")
    'aquiaquiaqui falta informar o numero abaixo
    lLinhaDados = lLinhaDados & Format(rsTabela("Contagem de Reinicio de Operacao").Value, "000") 'Contador de Reinicio de Operacao
    'lLinhaDados = lLinhaDados & Mid(Format(rsTabela("Totalizador Geral Inicial").Value, "00000000000000.00"), 1, 14) & Mid(Format(rsTabela("Totalizador Geral Inicial").Value, "00000000000000.00"), 16, 2)
    lLinhaDados = lLinhaDados & Mid(Format(rsTabela("Totalizador Geral Final").Value - rsTabela("Totalizador Geral Inicial").Value, "00000000000000.00"), 1, 14) & Mid(Format(rsTabela("Totalizador Geral Final").Value - rsTabela("Totalizador Geral Inicial").Value, "00000000000000.00"), 16, 2)
    lLinhaDados = lLinhaDados & Mid(Format(rsTabela("Totalizador Geral Final").Value, "00000000000000.00"), 1, 14) & Mid(Format(rsTabela("Totalizador Geral Final").Value, "00000000000000.00"), 16, 2)
    lLinhaDados = lLinhaDados & Space(37) '37 brancos
    
    'Grava String "lLinhaDados" no Arquivo Texto
    Print #1, lLinhaDados
    lRegistro = lRegistro + 1
    lRegistro60M = lRegistro60M + 1
    Exit Sub

ErrorRotina:
    Close #1
    MsgBox "Não foi possível gravar registro tipo 60M." & Chr(10) & "Erro de número " & Err, vbCritical, "GravaArquivoRegistro60M"
    Exit Sub
End Sub
Private Sub GravaArquivoRegistro71()
    Dim i As Integer
    Dim i2 As Integer
    Dim x_serie As String
    Dim xNumeroItem As Integer
    Dim xNumeroNF As Long
    Dim lCNPJ As String
    Dim xCodigo As String
    
    On Error GoTo ErrorRotina
    
    With rsItemEntradaSaida
        .Sort = "CNPJ, Serie, Numero, Item"
        .MoveFirst
        Do Until .EOF
            'If !numero = "3905" Then
            '   MsgBox !numero
            'End If
            If !CFOP = 2353 Then
                'Registro Tipo 71 (Item de NF Entrada)
'                tbl_movimento_cabecalho_nota_fiscal_entrada.Seek "=", g_empresa, CDate(!Data), !numero, !Serie
'                If tbl_movimento_cabecalho_nota_fiscal_entrada.NoMatch Then
'                    MsgBox "Entrada Inexistente!", vbInformation, "Erro de Integridade!"
'                    lLinhaDados = "71"
'                    lLinhaDados = lLinhaDados & !CNPJ
'                Else
'                    lIE = "              "
'                    tbl_fornecedor.Seek "=", g_empresa, tbl_movimento_cabecalho_nota_fiscal_entrada![Codigo do Fornecedor]
'                    If tbl_fornecedor.NoMatch Then
'                        MsgBox "Fornecedor não cadastrado!", vbCritical, "Erro na geração do disquete!"
'                    End If
'                    i2 = 0
'                    For i = 1 To 20
'                        If Mid(tbl_fornecedor![Inscricao Estadual], i, 1) >= "0" And Mid(tbl_fornecedor![Inscricao Estadual], i, 1) <= "9" Then
'                            i2 = i2 + 1
'                            Mid(lIE, i2, 1) = Mid(tbl_fornecedor![Inscricao Estadual], i, 1)
'                        End If
'                    Next
'                    lCNPJ = "              "
'                    Mid(lCNPJ, 1, 14) = tbl_fornecedor!CGC
'                    lLinhaDados = "71"
'                    lLinhaDados = lLinhaDados & lCNPJ
'                    lLinhaDados = lLinhaDados & lIE
'                    'lLinhaDados = lLinhaDados & lEmpresaCNPJ
'                    'lLinhaDados = lLinhaDados & lEmpresaIE
'                    lLinhaDados = lLinhaDados & Format(!Data, "yyyymmdd")
'                    'lLinhaDados = lLinhaDados & lEmpresaUF
'                    lLinhaDados = lLinhaDados & tbl_fornecedor!UF
'                    lLinhaDados = lLinhaDados & "08" '!Modelo
'                    lLinhaDados = lLinhaDados & !Serie
'                    lLinhaDados = lLinhaDados & !numero
'                    xNumeroNF = tbl_movimento_cabecalho_nota_fiscal_entrada![Numero da NF de Origem 1]
'                    'If xNumeroNF = "17473" Then
'                    '   MsgBox xNumeroNF
'                    'End If
'
'                    tbl_movimento_cabecalho_nota_fiscal_entrada.Seek ">=", g_empresa, CDate(!Data), xNumeroNF, " "
'                    If Not tbl_movimento_cabecalho_nota_fiscal_entrada.NoMatch Then
'                        If tbl_movimento_cabecalho_nota_fiscal_entrada!numero = xNumeroNF Then
'                            lIE = "              "
'                            tbl_fornecedor.Seek "=", g_empresa, tbl_movimento_cabecalho_nota_fiscal_entrada![Codigo do Fornecedor]
'                            If tbl_fornecedor.NoMatch Then
'                                MsgBox "Fornecedor não cadastrado!", vbCritical, "Erro na geração do disquete!"
'                            End If
'                            i2 = 0
'                            For i = 1 To 20
'                                If Mid(tbl_fornecedor![Inscricao Estadual], i, 1) >= "0" And Mid(tbl_fornecedor![Inscricao Estadual], i, 1) <= "9" Then
'                                    i2 = i2 + 1
'                                    Mid(lIE, i2, 1) = Mid(tbl_fornecedor![Inscricao Estadual], i, 1)
'                                End If
'                            Next
'                            lLinhaDados = lLinhaDados & tbl_fornecedor!UF
'                            lLinhaDados = lLinhaDados & !CNPJ
'                            lLinhaDados = lLinhaDados & lIE
'                            lLinhaDados = lLinhaDados & Format(tbl_movimento_cabecalho_nota_fiscal_entrada![Data de Entrada], "yyyymmdd")
'                            lLinhaDados = lLinhaDados & "01"
'                            x_serie = "  "
'                            Mid(x_serie, 1, 2) = tbl_movimento_cabecalho_nota_fiscal_entrada!Serie
'                            If x_serie = "U " Then
'                                lLinhaDados = lLinhaDados & "1  "
'                            Else
'                                lLinhaDados = lLinhaDados & x_serie & " "
'                            End If
'                            lLinhaDados = lLinhaDados & Format(tbl_movimento_cabecalho_nota_fiscal_entrada!numero, "000000")
'                            lLinhaDados = lLinhaDados & Mid(Format(tbl_movimento_cabecalho_nota_fiscal_entrada![Total da Nota], "000000000000.00"), 1, 12)
'                            lLinhaDados = lLinhaDados & Mid(Format(tbl_movimento_cabecalho_nota_fiscal_entrada![Total da Nota], "000000000000.00"), 14, 2)
'                            lLinhaDados = lLinhaDados & Space(12)
'                        Else
'                            MsgBox "NF com data inconsistente: " & xNumeroNF, vbInformation, "Erro de Consistência"
'                        End If
'                    Else
'                        MsgBox "TESTE"
'                    End If
'                End If
                Print #1, lLinhaDados
                lRegistro = lRegistro + 1
                lRegistro71 = lRegistro71 + 1
            End If
            .MoveNext
        Loop
    End With
    Exit Sub
ErrorRotina:
    Close #1
    'If Err = 53 Then
    '    MsgBox "Arquivo não encontrado para a data informada.", vbCritical, "Erro no processamento"
    '    Exit Sub
    'Else
        MsgBox "Não foi possível processar o disquete." & Chr(10) & "Erro de número " & Err, vbCritical, "GravaArquivoRegistro71"
        Exit Sub
    'End If
End Sub
Private Sub GravaArquivoRegistro75()
    Dim i As Integer
    Dim xNomeProduto As String
    Dim xUn As String
    Dim xCodigoAnterior As String
    Dim lAliquotaICMS As Currency
    Dim xCodigoProduto As String
    
    On Error GoTo ErrorRotina
    
    xCodigoAnterior = ""
    With rsItemEntradaSaida
        .Sort = "CodigoProduto"
        .MoveFirst
        Do Until .EOF
            If !CodigoProduto <> xCodigoAnterior Then
                xCodigoAnterior = !CodigoProduto
                'Registro Tipo 75
                'Registro de Código de Produto e Serviço
                lLinhaDados = "75"
                'Data Inicial
                lLinhaDados = lLinhaDados & Mid(msk_data_i.Text, 7, 4) & Mid(msk_data_i.Text, 4, 2) & Mid(msk_data_i.Text, 1, 2)
                'Data Final
                lLinhaDados = lLinhaDados & Mid(msk_data_f.Text, 7, 4) & Mid(msk_data_f.Text, 4, 2) & Mid(msk_data_f.Text, 1, 2)
                
                lAliquotaICMS = 0
                xCodigoProduto = Space(14)
                xNomeProduto = Space(53)
                xUn = Space(6)
                Mid(xCodigoProduto, 1, 14) = !CodigoProduto
                If Produto.LocalizarCodigo(CLng(Trim(xCodigoProduto))) Then
                    i = Len(Produto.Nome)
                    Mid(xNomeProduto, 1, i) = Produto.Nome
                    i = Len(Produto.Unidade)
                    Mid(xUn, 1, i) = Produto.Unidade
                    lAliquotaICMS = AliquotaItemICMS(Produto.CodigoAliquota)
                Else
                    MsgBox "Produto não cadastrado!" & vbCrLf & "Codigo=" & !CodigoProduto, vbCritical, "Err de Integridade!"
                End If
                lLinhaDados = lLinhaDados & xCodigoProduto
                'Codificação da Nomenclatura Comum do Mercosul
                lLinhaDados = lLinhaDados & "        "
                'Descrição do produto ou serviço
                lLinhaDados = lLinhaDados & xNomeProduto
                'Unidade de Medida de Comercialização
                lLinhaDados = lLinhaDados & xUn
                
                
                
                ''Código da situação tributária do produto ou serviço
                'lLinhaDados = lLinhaDados & "060"
                ''Alíquota do IPI
                'lLinhaDados = lLinhaDados & "0000"
                
                'Alíquota do IPI
                lLinhaDados = lLinhaDados & "00000"
                'Alíquota do ICMS
                lLinhaDados = lLinhaDados & Mid(Format(lAliquotaICMS, "00.00"), 1, 2) & Mid(Format(lAliquotaICMS, "00.00"), 4, 2)
                '% de Redução na base de cálculo do ICMS, nas operações internas
                lLinhaDados = lLinhaDados & "00000"
                'Base de Cálculo do ICMS de Substituição Tributária
                lLinhaDados = lLinhaDados & "0000000000000"
                If UCase(xNomeProduto) Like "*FRETE*" Then
                Else
                    Print #1, lLinhaDados
                    lRegistro = lRegistro + 1
                    lRegistro75 = lRegistro75 + 1
                End If
            End If
            .MoveNext
        Loop
    End With
    Exit Sub

ErrorRotina:
    Close #1
    'If Err = 53 Then
    '    MsgBox "Arquivo não encontrado para a data informada.", vbCritical, "Erro no processamento"
    '    Exit Sub
    'Else
        MsgBox "Não foi possível processar o arquivo." & Chr(10) & "Erro de número " & Err, vbCritical, "GravaArquivoRegistro75"
        Exit Sub
    'End If
End Sub
Private Sub GravaArquivoRegistro90()
    Dim i As Integer
    
    On Error GoTo ErrorRotina
    
    lRegistro = lRegistro + 1
    'Registro tipo 90
    'Totalização do Arquivo
    lLinhaDados = "90"
    'Numero do CNPJ
    lLinhaDados = lLinhaDados & Mid(lEmpresaCNPJ, 1, 14)
    'Inscricao Estadual
    lLinhaDados = lLinhaDados & lEmpresaIE
    'Tipo a Ser Totalizado e Total de Registro "50"
    lLinhaDados = lLinhaDados & "50" & Format(lRegistro50, "00000000")
    ''Tipo a Ser Totalizado e Total de Registro "53"
    'lLinhaDados = lLinhaDados & "53" & Format(lRegistro53, "00000000")
    'Tipo a Ser Totalizado e Total de Registro "54"
    lLinhaDados = lLinhaDados & "54" & Format(lRegistro54, "00000000")
    'Tipo a Ser Totalizado e Total de Registro "60" A + M
    lLinhaDados = lLinhaDados & "60" & Format(lRegistro60A + lRegistro60M, "00000000")
    'Tipo a Ser Totalizado e Total de Registro "70"
    lLinhaDados = lLinhaDados & "70" & Format(lRegistro70, "00000000")
    'Tipo a Ser Totalizado e Total de Registro "71"
    lLinhaDados = lLinhaDados & "71" & Format(lRegistro71, "00000000")
    'Tipo a Ser Totalizado e Total de Registro "75"
    lLinhaDados = lLinhaDados & "75" & Format(lRegistro75, "00000000")
    'Tipo a Ser Totalizado e Total de Registro "99"
    lRegistro50 = lRegistro50 + lRegistro53 + lRegistro54 + lRegistro60A + lRegistro60M + lRegistro70 + lRegistro71 + lRegistro75
    lLinhaDados = lLinhaDados & "99" & Format(lRegistro50 + 3, "00000000")
    lLinhaDados = lLinhaDados & String(25, " ")
    'Total de Registross
    'lLinhaDados = lLinhaDados & Format(lRegistro, "00000000")
    '8 espacos
    lLinhaDados = lLinhaDados & "1"
    Print #1, lLinhaDados
    Exit Sub

ErrorRotina:
    Close #1
    'If Err = 53 Then
    '    MsgBox "Arquivo não encontrado para a data informada.", vbCritical, "Erro no processamento"
    '    Exit Sub
    'Else
        MsgBox "Não foi possível processar o arquivo." & Chr(10) & "Erro de número " & Err, vbCritical, "GravaArquivoRegistro54"
        Exit Sub
    'End If
End Sub
Private Sub CriarsItemEntradaSaida()
    With rsItemEntradaSaida
        If lrsItemEntradaSaidaCriado Then
            .MoveFirst
            Do Until .EOF
                .Delete
                .MoveNext
            Loop
        Else
            .CursorLocation = adUseClient
            .Fields.Append "CNPJ", adVarChar, 14
            .Fields.Append "IE", adVarChar, 14
            .Fields.Append "DATA", adVarChar, 10
            .Fields.Append "Modelo", adVarChar, 2
            .Fields.Append "Serie", adVarChar, 3
            .Fields.Append "Numero", adVarChar, 6
            .Fields.Append "CFOP", adVarChar, 4
            .Fields.Append "CST", adVarChar, 3
            .Fields.Append "Item", adVarChar, 3
            .Fields.Append "CodigoProduto", adVarChar, 14
            .Fields.Append "Quantidade", adVarChar, 11
            .Fields.Append "ValorProduto", adVarChar, 12
            .Fields.Append "ValorDesconto", adVarChar, 12
            .Fields.Append "BaseICMS", adVarChar, 12
            .Fields.Append "BaseICMSST", adVarChar, 12
            .Fields.Append "ValorIPI", adVarChar, 12
            .Fields.Append "Aliquota", adVarChar, 4
            .Fields.Append "UF", adVarChar, 2
            .Open
            lrsItemEntradaSaidaCriado = True
        End If
    End With
End Sub
Private Sub GeraRsItemEntradaCombustivel()
    Dim xModelo As String
    Dim xSerie As String
    Dim xNumeroNF As String
    Dim xItem As Integer
    Dim xCodigo As String
    
    On Error GoTo ErrorRotina
    
    
    '********************************
    '**** ENTRADA DE COMBUSTIVEL ****
    '********************************
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Data, [Tipo de Combustivel], [Numero da Nota], [Codigo do Fornecedor], "
    lSQL = lSQL & "       [Valor do Litro], Quantidade, [Valor da Entrada], Modelo, Serie"
    lSQL = lSQL & "  FROM Entrada_Combustivel_LMC"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND Data <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & " ORDER BY Data, [Numero da Nota], [Codigo do Fornecedor]"

    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    
    If rsTabela.RecordCount > 0 Then
        rsTabela.MoveFirst
        Do Until rsTabela.EOF
            lCNPJ = "00000000000000"
            If Not Fornecedor.LocalizarCodigo(g_empresa, rsTabela("Codigo do Fornecedor").Value) Then
                MsgBox "Fornecedor não cadastrado!", vbCritical, "Erro na geração do arquivo!"
                Exit Sub
            Else
                
            Mid(lCNPJ, 1, 14) = Fornecedor.CGC
                lIE = DesmascaraIscricaoEstadual(Fornecedor.InscricaoEstadual)
                lUF = Fornecedor.UF
            End If
            
            If xNumeroNF <> rsTabela("Numero da Nota").Value Then
                xNumeroNF = rsTabela("Numero da Nota").Value
                xItem = 0
            End If
            xItem = xItem + 1
            
            rsItemEntradaSaida.AddNew
            rsItemEntradaSaida!CNPJ = lCNPJ
            rsItemEntradaSaida!IE = lIE
            rsItemEntradaSaida!Data = Format(rsTabela("Data").Value, "dd/mm/yyyy")
            xModelo = Space(2)
            Mid(xModelo, 1, 2) = rsTabela("Modelo").Value
            rsItemEntradaSaida!Modelo = xModelo
            xSerie = Space(3)
            Mid(xSerie, 1, 3) = rsTabela("Serie").Value
            rsItemEntradaSaida!Serie = xSerie
            rsItemEntradaSaida!numero = Format(rsTabela("Numero da Nota").Value, "000000")
            
            rsItemEntradaSaida!CFOP = "1652"
            'ligar nba contadora do givanildo e pegar codigo
            
            rsItemEntradaSaida!CST = "000"
            rsItemEntradaSaida!Item = Format(xItem, "000")
            xCodigo = Space(14)
            'i = Len(rsTabela("Tipo de Combustivel").Value)
            'Mid(xCodigo, 1, i) = rsTabela("Tipo de Combustivel").Value
            If IsNumeric(Trim(rsTabela("Tipo de Combustivel").Value)) Then
                Mid(xCodigo, 1, 14) = rsTabela("Tipo de Combustivel").Value
            Else
                If Bomba.LocalizarTipoCombustivel(g_empresa, rsTabela("Tipo de Combustivel").Value) Then
                    Mid(xCodigo, 1, 14) = Bomba.CodigoProduto
                End If
            End If
            rsItemEntradaSaida!CodigoProduto = xCodigo
            rsItemEntradaSaida!Quantidade = Mid(Format(rsTabela("Quantidade").Value, "00000000.000"), 1, 8) & Mid(Format(rsTabela("Quantidade").Value, "00000000.000"), 10, 3)
            rsItemEntradaSaida!ValorProduto = Mid(Format(rsTabela("Valor da Entrada").Value, "0000000000.00"), 1, 10) & Mid(Format(rsTabela("Valor da Entrada").Value, "0000000000.00"), 12, 2)
            rsItemEntradaSaida!ValorDesconto = Mid(Format(0, "0000000000.00"), 1, 10) & Mid(Format(0, "0000000000.00"), 12, 2)
            
            'If rsTabela("Tipo de Combustivel").Value = "A " Then
            'Else
            'End If
            rsItemEntradaSaida!BaseICMS = "000000000000"
            rsItemEntradaSaida!BaseICMSST = Mid(Format(0, "0000000000.00"), 1, 10) & Mid(Format(0, "0000000000.00"), 12, 2)
            rsItemEntradaSaida!ValorIPI = Mid(Format(0, "0000000000.00"), 1, 10) & Mid(Format(0, "0000000000.00"), 12, 2)
            rsItemEntradaSaida!Aliquota = Mid(Format(0, "00.00"), 1, 2) & Mid(Format(0, "00.00"), 4, 2)
            rsItemEntradaSaida!UF = lUF
            rsItemEntradaSaida.Update
            rsTabela.MoveNext
        Loop
    End If
    rsTabela.Close
    Set rsTabela = Nothing
    Exit Sub

ErrorRotina:
    Close #1
    MsgBox "Não foi possível processar o arquivo." & Chr(10) & "Erro de número " & Err, vbCritical, "GeraRsItemEntradaCombustivel"
    Exit Sub
End Sub
Private Sub GeraRsItemEntradaProduto()
    Dim xModelo As String
    Dim xSerie As String
    Dim xNumeroNF As String
    Dim xItem As Integer
    Dim xCodigo As String
    
    On Error GoTo ErrorRotina
    
    '****************************
    '**** ENTRADA DE PRODUTO ****
    '****************************
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT [Data da Entrada], [Codigo do Produto], [Numero do Documento], [Codigo do Fornecedor], "
    lSQL = lSQL & "       Quantidade, [Total do Custo], CFOP, Modelo, Serie"
    lSQL = lSQL & "  FROM Entrada_Produto"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND [Data da Entrada] >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND [Data da Entrada] <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & "   AND [Tipo da Entrada] = " & preparaTexto("1")
    lSQL = lSQL & " ORDER BY [Data da Entrada], [Numero do Documento], [Codigo do Fornecedor]"

    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    
    If rsTabela.RecordCount > 0 Then
        rsTabela.MoveFirst
        Do Until rsTabela.EOF
            lCNPJ = "00000000000000"
            If Not Fornecedor.LocalizarCodigo(g_empresa, rsTabela("Codigo do Fornecedor").Value) Then
                MsgBox "Fornecedor não cadastrado!", vbCritical, "Erro na geração do arquivo!"
                Exit Sub
            Else
                
            Mid(lCNPJ, 1, 14) = Fornecedor.CGC
                lIE = DesmascaraIscricaoEstadual(Fornecedor.InscricaoEstadual)
                lUF = Fornecedor.UF
            End If
            
            If xNumeroNF <> rsTabela("Numero do Documento").Value Then
                xNumeroNF = rsTabela("Numero do Documento").Value
                xItem = 0
            End If
            xItem = xItem + 1
            
            rsItemEntradaSaida.AddNew
            rsItemEntradaSaida!CNPJ = lCNPJ
            rsItemEntradaSaida!IE = lIE
            rsItemEntradaSaida!Data = Format(rsTabela("Data da Entrada").Value, "dd/mm/yyyy")
            xModelo = Space(2)
            Mid(xModelo, 1, 2) = rsTabela("Modelo").Value
            rsItemEntradaSaida!Modelo = xModelo
            xSerie = Space(3)
            Mid(xSerie, 1, 3) = rsTabela("Serie").Value
            rsItemEntradaSaida!Serie = xSerie
            rsItemEntradaSaida!numero = Format(CLng(rsTabela("Numero do Documento").Value), "000000")
            rsItemEntradaSaida!CFOP = rsTabela("CFOP").Value
            
            
            'ligar nba contadora do givanildo e pegar codigo
            
            rsItemEntradaSaida!CST = "000"
            rsItemEntradaSaida!Item = Format(xItem, "000")
            xCodigo = Space(14)
            Mid(xCodigo, 1, 14) = rsTabela("Codigo do Produto").Value
            rsItemEntradaSaida!CodigoProduto = xCodigo
            rsItemEntradaSaida!Quantidade = Mid(Format(rsTabela("Quantidade").Value, "00000000.000"), 1, 8) & Mid(Format(rsTabela("Quantidade").Value, "00000000.000"), 10, 3)
            rsItemEntradaSaida!ValorProduto = Mid(Format(rsTabela("Total do Custo").Value, "0000000000.00"), 1, 10) & Mid(Format(rsTabela("Total do Custo").Value, "0000000000.00"), 12, 2)
            rsItemEntradaSaida!ValorDesconto = Mid(Format(0, "0000000000.00"), 1, 10) & Mid(Format(0, "0000000000.00"), 12, 2)
            
            rsItemEntradaSaida!BaseICMS = "000000000000"
            rsItemEntradaSaida!BaseICMSST = Mid(Format(0, "0000000000.00"), 1, 10) & Mid(Format(0, "0000000000.00"), 12, 2)
            rsItemEntradaSaida!ValorIPI = Mid(Format(0, "0000000000.00"), 1, 10) & Mid(Format(0, "0000000000.00"), 12, 2)
            rsItemEntradaSaida!Aliquota = Mid(Format(0, "00.00"), 1, 2) & Mid(Format(0, "00.00"), 4, 2)
            rsItemEntradaSaida!UF = lUF
            rsItemEntradaSaida.Update
            rsTabela.MoveNext
        Loop
    End If
    rsTabela.Close
    Set rsTabela = Nothing
    Exit Sub

ErrorRotina:
    Close #1
    MsgBox "Não foi possível processar o arquivo." & Chr(10) & "Erro de número " & Err, vbCritical, "GeraRsItemEntradaProduto"
    Exit Sub
End Sub
Private Sub GeraRsItemSaida()
    Dim i As Integer
    Dim x_serie As String * 2
    Dim xNumeroItem As Integer
    Dim lCNPJ As String
    Dim xCodigo As String
    Dim lAliquotaICMS As Currency
    Dim xReducaoBC As Boolean
    Dim xBaseCalculoICMS As Currency
    Dim xDescontoEfetuado As Integer
    
    On Error GoTo ErrorRotina
    
    
    '********************************
    '**** ENTRADA DE PRODUTOS    ****
    '********************************
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Data, [Tipo de Combustivel], [Numero da Nota], [Valor do Litro], "
    lSQL = lSQL & "       Quantidade, [Valor da Entrada], [Codigo do Fornecedor], [Tipo de Transporte]"
    lSQL = lSQL & "  FROM Entrada_Combustivel_LMC"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND Data <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & " ORDER BY Data, [Numero da Nota]"

    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    
    If rsTabela.RecordCount > 0 Then
        rsTabela.MoveFirst
        Do Until rsTabela.EOF
            'rsTabela("Nome").Value
           
            
            rsTabela.MoveNext
        Loop
    End If
    rsTabela.Close
    Set rsTabela = Nothing

    
    
    
    
    
    
    
    
'    With tbl_movimento_cabecalho_nota_fiscal_entrada
'        .Seek ">=", g_empresa, CDate(msk_data_i), 0, "  "
'        If Not .NoMatch Then
'            Do Until .EOF
'                If !Empresa <> g_empresa Or ![Data de Entrada] > CDate(msk_data_f) Then
'                    Exit Do
'                End If
'                If !Serie <> "AC" And !Serie <> "CP" Then
'                    'Registro Tipo 54 (Entradas)
'                    'Produto
'                    xDescontoEfetuado = 0
'                    xNumeroItem = 0
'                    tbl_movimento_nota_fiscal_entrada.Seek ">=", g_empresa, ![Data de Entrada], !numero, !Serie, 0
'                    If Not tbl_movimento_nota_fiscal_entrada.NoMatch Then
'                        Do Until tbl_movimento_nota_fiscal_entrada.EOF
'
'                            'If tbl_movimento_nota_fiscal_entrada!numero = 510732 Then
'                            '    MsgBox tbl_movimento_nota_fiscal_entrada!numero
'                            'End If
'
'
'                            If tbl_movimento_nota_fiscal_entrada!Empresa <> g_empresa Or tbl_movimento_nota_fiscal_entrada![Data de Entrada] <> ![Data de Entrada] Or tbl_movimento_nota_fiscal_entrada!numero <> !numero Or tbl_movimento_nota_fiscal_entrada!Serie <> !Serie Then
'                                Exit Do
'                            End If
'                            xNumeroItem = xNumeroItem + 1
'                            lCNPJ = "00000000000000"
'                            If ![Codigo do Fornecedor] > 0 Then
'                                tbl_fornecedor.Seek "=", g_empresa, ![Codigo do Fornecedor]
'                                If tbl_fornecedor.NoMatch Then
'                                    MsgBox "Fornecedor não cadastrado!", vbCritical, "Erro na geração do disquete!"
'                                    Exit Sub
'                                End If
'                                lCNPJ = tbl_fornecedor!CGC
'                            End If
'                            rsItemEntradaSaida.AddNew
'                            rsItemEntradaSaida!CNPJ = lCNPJ
'                            rsItemEntradaSaida!IE = lIE
'                            rsItemEntradaSaida!Data = Format(![Data de Entrada], "dd/mm/yyyy")
'                            rsItemEntradaSaida!Modelo = "01"
'                            x_serie = !Serie
'                            rsItemEntradaSaida!Serie = x_serie & " "
'                            rsItemEntradaSaida!numero = Format(!numero, "000000")
'                            'rsItemEntradaSaida!CFOP = ![Codificacao Fiscal]
'                            rsItemEntradaSaida!CFOP = tbl_movimento_nota_fiscal_entrada!CFOP
'                            'If rsItemEntradaSaida!CFOP = "2353" Then
'                            '    rsItemEntradaSaida!CFOP = "2102"
'                            'End If
'                            rsItemEntradaSaida!CST = "000"
'                            rsItemEntradaSaida!Item = Format(xNumeroItem, "000")
'                            xCodigo = Space(14)
'                            i = Len(tbl_movimento_nota_fiscal_entrada![Codigo do Produto])
'                            Mid(xCodigo, 1, i) = tbl_movimento_nota_fiscal_entrada![Codigo do Produto]
'                            rsItemEntradaSaida!CodigoProduto = xCodigo
'                            rsItemEntradaSaida!Quantidade = Mid(Format(tbl_movimento_nota_fiscal_entrada!Quantidade, "00000000.000"), 1, 8) & Mid(Format(tbl_movimento_nota_fiscal_entrada!Quantidade, "00000000.000"), 10, 3)
'                            rsItemEntradaSaida!ValorProduto = Mid(Format(tbl_movimento_nota_fiscal_entrada![Valor Total], "0000000000.00"), 1, 10) & Mid(Format(tbl_movimento_nota_fiscal_entrada![Valor Total], "0000000000.00"), 12, 2)
'                            rsItemEntradaSaida!ValorDesconto = Mid(Format(0, "0000000000.00"), 1, 10) & Mid(Format(0, "0000000000.00"), 12, 2)
'                            If ![Valor do Desconto] < tbl_movimento_nota_fiscal_entrada![Valor Total] And xDescontoEfetuado = 0 Then
'                                rsItemEntradaSaida!ValorDesconto = Mid(Format(![Valor do Desconto], "0000000000.00"), 1, 10) & Mid(Format(![Valor do Desconto], "0000000000.00"), 12, 2)
'                                xDescontoEfetuado = 1
'                            End If
'                            If tbl_movimento_nota_fiscal_entrada![Aliquota de ICMS] > 0 Then
'                                If xDescontoEfetuado = 1 Then
'                                    xDescontoEfetuado = 2
'                                    If tbl_movimento_nota_fiscal_entrada![Valor Total] < ![Valor de Outros] Then
'                                        rsItemEntradaSaida!BaseICMS = "000000000000"
'                                    Else
'                                        rsItemEntradaSaida!BaseICMS = Mid(Format(tbl_movimento_nota_fiscal_entrada![Valor Total] - ![Valor do Desconto] - ![Valor de Outros] + tbl_movimento_nota_fiscal_entrada![Valor do IPI], "0000000000.00"), 1, 10) & Mid(Format(tbl_movimento_nota_fiscal_entrada![Valor Total] - ![Valor do Desconto] - ![Valor de Outros] + tbl_movimento_nota_fiscal_entrada![Valor do IPI], "0000000000.00"), 12, 2)
'                                    End If
'                                Else
'                                    rsItemEntradaSaida!BaseICMS = Mid(Format(tbl_movimento_nota_fiscal_entrada![Valor Total], "0000000000.00"), 1, 10) & Mid(Format(tbl_movimento_nota_fiscal_entrada![Valor Total], "0000000000.00"), 12, 2)
'                                End If
'                            Else
'                                rsItemEntradaSaida!BaseICMS = Mid(Format(0, "0000000000.00"), 1, 10) & Mid(Format(0, "0000000000.00"), 12, 2)
'                            End If
'                            rsItemEntradaSaida!BaseICMSST = Mid(Format(0, "0000000000.00"), 1, 10) & Mid(Format(0, "0000000000.00"), 12, 2)
'                            rsItemEntradaSaida!ValorIPI = Mid(Format(tbl_movimento_nota_fiscal_entrada![Valor do IPI], "0000000000.00"), 1, 10) & Mid(Format(tbl_movimento_nota_fiscal_entrada![Valor do IPI], "0000000000.00"), 12, 2)
'                            rsItemEntradaSaida!Aliquota = Mid(Format(tbl_movimento_nota_fiscal_entrada![Aliquota de ICMS], "00.00"), 1, 2) & Mid(Format(tbl_movimento_nota_fiscal_entrada![Aliquota de ICMS], "00.00"), 4, 2)
'                            rsItemEntradaSaida.Update
'                            tbl_movimento_nota_fiscal_entrada.MoveNext
'                        Loop
'                    End If
'                End If
'                .MoveNext
'            Loop
'        End If
'    End With




'    With tbl_movimento_cabecalho_nota_fiscal_saida
'        .Seek ">=", g_empresa, CDate(msk_data_i), 0, 0, "  "
'        If Not .NoMatch Then
'            Do Until .EOF
'                If !Empresa <> g_empresa Or !Data > CDate(msk_data_f) Then
'                    Exit Do
'                End If
'                'If !numero = 713 Then
'                '    MsgBox !numero
'                'End If
'                xReducaoBC = False
'                If !Serie <> "AC" And !Serie <> "CP" And ![Situacao da Venda] <= 2 Then
'                    If !Cancelada = False Then
'                        'Registro Tipo 54 (Saidas)
'                        'Produto
'                        xNumeroItem = 0
'                        tbl_movimento_nota_fiscal_saida.Seek ">=", g_empresa, !Data, !Impressora, !numero, !Serie, 0
'                        If Not tbl_movimento_nota_fiscal_saida.NoMatch Then
'                            Do Until tbl_movimento_nota_fiscal_saida.EOF
'                                If tbl_movimento_nota_fiscal_saida!Empresa <> g_empresa Or tbl_movimento_nota_fiscal_saida!Data <> !Data Or tbl_movimento_nota_fiscal_saida!Impressora <> !Impressora Or tbl_movimento_nota_fiscal_saida!numero <> !numero Or tbl_movimento_nota_fiscal_saida!Serie <> !Serie Then
'                                    Exit Do
'                                End If
'                                xNumeroItem = xNumeroItem + 1
'                                lCNPJ = "00000000000000"
'                                If ![Codigo do Cliente] > 0 Then
'                                    tbl_cliente.Seek "=", ![Codigo do Cliente]
'                                    If Not tbl_cliente.NoMatch Then
'                                        If Len(Trim(tbl_cliente!CGC)) > 0 Then
'                                            lCNPJ = tbl_cliente!CGC
'                                        ElseIf Len(Trim(tbl_cliente!CPF)) > 0 Then
'                                            lCNPJ = "000" & tbl_cliente!CPF
'                                        End If
'                                        If tbl_cliente!UF = g_uf_empresa Then
'                                            If Val(tbl_cliente![Inscricao Estadual]) > 0 Then
'                                                xReducaoBC = True
'                                            End If
'                                        End If
'                                    End If
'                                End If
'                                rsItemEntradaSaida.AddNew
'                                rsItemEntradaSaida!CNPJ = lCNPJ
'                                rsItemEntradaSaida!IE = lIE
'                                rsItemEntradaSaida!Data = Format(!Data, "dd/mm/yyyy")
'                                rsItemEntradaSaida!Modelo = "01"
'                                x_serie = !Serie
'                                rsItemEntradaSaida!Serie = x_serie & " "
'                                rsItemEntradaSaida!numero = Format(!numero, "000000")
'                                rsItemEntradaSaida!CFOP = ![Codificacao Fiscal]
'                                rsItemEntradaSaida!CST = "000"
'                                rsItemEntradaSaida!Item = Format(xNumeroItem, "000")
'                                xCodigo = Space(14)
'                                i = Len(tbl_movimento_nota_fiscal_saida![Codigo do Produto])
'                                Mid(xCodigo, 1, i) = tbl_movimento_nota_fiscal_saida![Codigo do Produto]
'                                rsItemEntradaSaida!CodigoProduto = xCodigo
'                                rsItemEntradaSaida!Quantidade = Mid(Format(tbl_movimento_nota_fiscal_saida!Quantidade, "00000000.000"), 1, 8) & Mid(Format(tbl_movimento_nota_fiscal_saida!Quantidade, "00000000.000"), 10, 3)
'                                rsItemEntradaSaida!ValorProduto = Mid(Format(tbl_movimento_nota_fiscal_saida![Valor Total], "0000000000.00"), 1, 10) & Mid(Format(tbl_movimento_nota_fiscal_saida![Valor Total], "0000000000.00"), 12, 2)
'                                rsItemEntradaSaida!ValorDesconto = Mid(Format(0, "0000000000.00"), 1, 10) & Mid(Format(0, "0000000000.00"), 12, 2)
'                                xBaseCalculoICMS = 0
'                                If tbl_movimento_nota_fiscal_saida![Aliquota de ICMS] > 0 Then
'                                    If xReducaoBC Then
'                                        If tbl_movimento_nota_fiscal_saida![Codigo da Aliquota] <> 7 Then
'                                            xBaseCalculoICMS = Format(tbl_movimento_nota_fiscal_saida![Valor Total] - tbl_movimento_nota_fiscal_saida![Valor Total] * 16.67 / 100, "0000000000.00")
'                                        End If
'                                    Else
'                                        xBaseCalculoICMS = tbl_movimento_nota_fiscal_saida![Valor Total]
'                                    End If
'                                End If
'                                'rsItemEntradaSaida!BaseICMS = Mid(Format(0, "0000000000.00"), 1, 10) & Mid(Format(0, "0000000000.00"), 12, 2)
'                                'If tbl_movimento_nota_fiscal_saida![Aliquota de ICMS] > 0 Then
'                                '    rsItemEntradaSaida!BaseICMS = Mid(Format(tbl_movimento_nota_fiscal_saida![Valor Total], "0000000000.00"), 1, 10) & Mid(Format(tbl_movimento_nota_fiscal_saida![Valor Total], "0000000000.00"), 12, 2)
'                                'Else
'                                    rsItemEntradaSaida!BaseICMS = Mid(Format(xBaseCalculoICMS, "0000000000.00"), 1, 10) & Mid(Format(xBaseCalculoICMS, "0000000000.00"), 12, 2)
'                                'End If
'                                rsItemEntradaSaida!BaseICMSST = Mid(Format(0, "0000000000.00"), 1, 10) & Mid(Format(0, "0000000000.00"), 12, 2)
'                                rsItemEntradaSaida!ValorIPI = Mid(Format(0, "0000000000.00"), 1, 10) & Mid(Format(0, "0000000000.00"), 12, 2)
'                                'lAliquotaICMS = AliquotaItemICMS(tbl_movimento_nota_fiscal_saida![Codigo da Aliquota])
'                                lAliquotaICMS = tbl_movimento_nota_fiscal_saida![Aliquota de ICMS]
'                                lLinhaDados = lLinhaDados & Mid(Format(lAliquotaICMS, "00.00"), 1, 2) & Mid(Format(lAliquotaICMS, "00.00"), 4, 2)
'                                rsItemEntradaSaida!Aliquota = Mid(Format(lAliquotaICMS, "00.00"), 1, 2) & Mid(Format(lAliquotaICMS, "00.00"), 4, 2)
'                                rsItemEntradaSaida.Update
'                                tbl_movimento_nota_fiscal_saida.MoveNext
'                            Loop
'                        End If
'                    End If
'                End If
'                .MoveNext
'            Loop
'        End If
'    End With
    Exit Sub
ErrorRotina:
    Close #1
    MsgBox "Não foi possível processar o arquivo." & Chr(10) & "Erro de número " & Err, vbCritical, "GeraRsItemSaida"
    Exit Sub
End Sub
Private Function DesmascaraIscricaoEstadual(ByVal pIE As String) As String
    Dim i As Integer
    Dim i2 As Integer
    Dim xIE As String
    
    xIE = Space(14)
    For i = 1 To 20
        If Mid(pIE, i, 1) >= "0" And Mid(pIE, i, 1) <= "9" Then
            i2 = i2 + 1
            If i2 > 14 Then
                Exit For
            End If
            Mid(xIE, i2, 1) = Mid(pIE, i, 1)
        End If
    Next
    If Trim(xIE) = "" Then
        Mid(xIE, 1, 6) = "ISENTO"
    End If
    DesmascaraIscricaoEstadual = xIE
End Function
Private Sub Finaliza()
    Set Aliquota = Nothing
    Set Bomba = Nothing
    Set ECF = Nothing
    Set Empresa = Nothing
    Set Fornecedor = Nothing
    Set Produto = Nothing
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
    g_string = ""
    cmd_ok.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cmd_ok.SetFocus
    Else
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_ok_Click()
    If ValidaCampos Then
        LimpaArquivosTemporarioValidadorSintegra
        GeraArquivo
        PreparaBotoes
        cmd_sair.SetFocus
    End If
End Sub
Function QuantidadeAliquotaNFSaida(x_data As Date, x_impressora As Integer, x_numero As Long, x_serie As String) As Integer
    Dim i As Integer
    Dim xPercentualAliquota(1 To 6) As String
    QuantidadeAliquotaNFSaida = 0
'    With tbl_movimento_nota_fiscal_saida
'        .Seek ">=", g_empresa, x_data, x_impressora, x_numero, x_serie, 0
'        If Not .NoMatch Then
'            Do Until .EOF
'                If !Empresa <> g_empresa Or !Data <> x_data Or !Impressora <> x_impressora Or !numero <> x_numero Or !Serie <> x_serie Then
'                    Exit Do
'                End If
'                For i = 1 To 6
'                    If xPercentualAliquota(i) = "" Then
'                        xPercentualAliquota(i) = ![Aliquota de ICMS]
'                        QuantidadeAliquotaNFSaida = QuantidadeAliquotaNFSaida + 1
'                        Exit For
'                    ElseIf xPercentualAliquota(i) = ![Aliquota de ICMS] Then
'                        Exit For
'                    End If
'                Next
'                .MoveNext
'            Loop
'        End If
'    End With
End Function
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data_i.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f.Text) < CDate(msk_data_i.Text) Then
        MsgBox "Data final deve ser maior ou igual a " & msk_data_i.Text & ".", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf Not Empresa.LocalizarCodigo(g_empresa) Then
        MsgBox "Empresa não cadastrada!" & vbCrLf & "Empresa = " & g_empresa, vbInformation, "Atenção!"
        cmd_sair.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Function ValorICMS(x_data As Date, x_impressora As Integer, x_numero As Long, x_serie As String) As Currency
    'loop tabela Movimento_Nota_Fiscal_Saida
    ValorICMS = 0
'    With tbl_movimento_nota_fiscal_saida
'        .Seek ">=", g_empresa, x_data, x_impressora, x_numero, x_serie, 0
'        If Not .NoMatch Then
'            Do Until .EOF
'                If !Empresa <> g_empresa Or !Data <> x_data Or !Impressora <> x_impressora Or !numero <> x_numero Or !Serie <> x_serie Then
'                    Exit Do
'                End If
'                'tbl_aliquota.Seek "=", ![Codigo da Aliquota]
'                'If Not tbl_aliquota.NoMatch Then
'                '    If tbl_aliquota![Aliquota do Imposto] = 0 Then
'                '        Exit Do
'                '    End If
'                '    ValorICMS = ValorICMS + Mid(Format((tbl_movimento_cabecalho_nota_fiscal_saida![Total da Nota] * tbl_aliquota![Aliquota do Imposto]) / 100, "0000000000.0000"), 1, 13)
'                '    Exit Function
'                'End If
'                ValorICMS = ValorICMS + Mid(Format((tbl_movimento_cabecalho_nota_fiscal_saida![Base de Calculo do ICMS] * ![Aliquota de ICMS]) / 100, "0000000000.0000"), 1, 13)
'                Exit Do
'                .MoveNext
'            Loop
'        End If
'    End With
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    PreparaBotoes
    msk_data_i.SetFocus
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    'Set tbl_aliquota = bd_sgle.OpenTable("Aliquota")
    'Set tbl_cliente = bd_sgle.OpenTable("Cliente")
    'Set tbl_empresa = bd_sgle.OpenTable("Empresas")
    'Set tbl_fornecedor = bd_sgle.OpenTable("Fornecedor")
    'Set tbl_movimento_cabecalho_nota_fiscal_entrada = bd_sgle.OpenTable("Movimento_Cabecalho_Nota_Fiscal_Entrada")
    'Set tbl_movimento_cabecalho_nota_fiscal_saida = bd_sgle.OpenTable("Movimento_Cabecalho_Nota_Fiscal_Saida")
    'Set tbl_movimento_nota_fiscal_entrada = bd_sgle.OpenTable("Movimento_Nota_Fiscal_Entrada")
    'Set tbl_movimento_nota_fiscal_saida = bd_sgle.OpenTable("Movimento_Nota_Fiscal_Saida")
    'Set tbl_produto = bd_sgle.OpenTable("Produto")
    'tbl_aliquota.Index = "id_codigo"
    'tbl_cliente.Index = "id_codigo"
    'tbl_empresa.Index = "id_codigo"
    'tbl_fornecedor.Index = "id_codigo"
    'tbl_movimento_cabecalho_nota_fiscal_entrada.Index = "id_data_entrada"
    'tbl_movimento_cabecalho_nota_fiscal_saida.Index = "id_data"
    'tbl_movimento_nota_fiscal_entrada.Index = "id_data"
    'tbl_movimento_nota_fiscal_saida.Index = "id_data"
    'tbl_produto.Index = "id_codigo2"
    lrsItemEntradaSaidaCriado = False
    msk_data_i.Text = fDataPrimeiroDiaMesAnterior(Date)
    msk_data_f.Text = fDataUltimoDiaMesAnterior(Date)
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
        cmd_ok.SetFocus
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
