VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_contas_pagar_especial 
   Caption         =   "Relação de Contas à Pagar Especial (Epson)"
   ClientHeight    =   3300
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_contas_pagar_especial.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_contas_pagar_especial.frx":030A
   ScaleHeight     =   3300
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_contas_pagar_especial.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Visualiza contas à pagar especial (Epson)."
      Top             =   2340
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_contas_pagar_especial.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprime contas à pagar especial (Epson)."
      Top             =   2340
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_contas_pagar_especial.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2340
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2115
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CheckBox chkUnificaEmpresa 
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   1740
         Width           =   435
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "lst_contas_pagar_especial.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_contas_pagar_especial.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_contas_pagar_especial.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbo_fornecedor 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1320
         Width           =   4755
      End
      Begin VB.ComboBox cbo_conta 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   960
         Width           =   3435
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4860
         TabIndex        =   8
         Top             =   600
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
         Top             =   600
         Width           =   1035
         _ExtentX        =   1826
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
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "&Unifica empresas"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1740
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "&Fornecedor"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Conta"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "&Data Inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata Final"
         Height          =   315
         Left            =   3840
         TabIndex        =   7
         Top             =   600
         Width           =   975
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
      Left            =   120
      Top             =   2340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_contas_pagar_especial"
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
Dim l_total As Currency
Dim lTotal(1 To 17, 1 To 11) As Currency
Dim lData(1 To 10) As Date
Dim lCodigoEmpresa(1 To 16) As Integer
Dim lNomeEmpresa(1 To 16) As String
Dim lSQL As String
Dim lDataSubTotal As Date
Dim lSubTotal As Currency

Private GrupoEmpresa As New cGrupoEmpresa
Private LocalCobranca As New cLocalCobranca
Private MovimentoContaPagar As New cMovimentoContaPagar
Private rstMovimentoContasPagar As New adodb.Recordset
Private Sub CalculaTotalEmpresas()
    Dim x_empresa As Integer
    Dim i As Integer
    
    lSQL = ""
    lSQL = lSQL & "SELECT Empresa, Data_Vencimento, SUM(Valor) AS Total"
    lSQL = lSQL & "  FROM Contas_Pagar"
    lSQL = lSQL & " WHERE Data_Vencimento >= " & preparaData(msk_data.Text)
    lSQL = lSQL & "   AND Data_Vencimento <= " & preparaData(CDate(msk_data.Text) + 9)
    If Val(cbo_conta.ItemData(cbo_conta.ListIndex)) > 0 Then
        lSQL = lSQL & "   AND Codigo_Conta = " & Val(cbo_conta.ItemData(cbo_conta.ListIndex))
    End If
    If Val(cbo_fornecedor.ItemData(cbo_fornecedor.ListIndex)) > 0 Then
        lSQL = lSQL & "   AND Codigo_Fornecedor = " & Val(cbo_fornecedor.ItemData(cbo_fornecedor.ListIndex))
    End If
    lSQL = lSQL & " GROUP BY Empresa, Data_Vencimento"
    lSQL = lSQL & " ORDER BY Empresa, Data_Vencimento"
    
    Set rstMovimentoContasPagar = Conectar.RsConexao(lSQL)
    
    With rstMovimentoContasPagar
        If .RecordCount > 0 Then
            For x_empresa = 1 To 16
                For i = 1 To 10
                    '.Seek ">=", lCodigoEmpresa(x_empresa), lData(i), 0
                    'If Not .NoMatch Then
                        Do Until .EOF
                            If !Empresa <> lCodigoEmpresa(x_empresa) Or !Data_Vencimento <> lData(i) Then
                                Exit Do
                            End If
                            lTotal(x_empresa, i) = lTotal(x_empresa, i) + !total
                            lTotal(17, i) = lTotal(17, i) + !total
                            lTotal(x_empresa, 11) = lTotal(x_empresa, 11) + !total
                            lTotal(17, 11) = lTotal(17, 11) + !total
                            .MoveNext
                        Loop
                    'End If
                Next
            Next
        End If
    End With
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    g_string = ""
    Set GrupoEmpresa = Nothing
    Set LocalCobranca = Nothing
    Set MovimentoContaPagar = Nothing
End Sub
Private Sub LoopMovimentoContasPagar()
    Dim x_local_cobranca As String
    
    
    lSQL = ""
    lSQL = lSQL & "SELECT Empresa, Data_Vencimento, Numero_Documento, Local_Cobranca, Data_Emissao, Valor, Nome_Fornecedor, Complemento, codigo_conta, codigo_fornecedor"
    lSQL = lSQL & "  FROM Contas_Pagar"
    lSQL = lSQL & " WHERE Data_Vencimento >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND Data_Vencimento <= " & preparaData(msk_data_f.Text)
    If chkUnificaEmpresa.Value = 0 Then
        lSQL = lSQL & "    AND Empresa = " & g_empresa
    End If
    If Val(cbo_conta.ItemData(cbo_conta.ListIndex)) > 0 Then
        lSQL = lSQL & "   AND Codigo_Conta = " & Val(cbo_conta.ItemData(cbo_conta.ListIndex))
    End If
    If Val(cbo_fornecedor.ItemData(cbo_fornecedor.ListIndex)) > 0 Then
        lSQL = lSQL & "   AND Codigo_Fornecedor = " & Val(cbo_fornecedor.ItemData(cbo_fornecedor.ListIndex))
    End If
    lSQL = lSQL & " ORDER BY Data_Vencimento, Data_Emissao"
    
    Set rstMovimentoContasPagar = Conectar.RsConexao(lSQL)
    With rstMovimentoContasPagar
        If .RecordCount > 0 Then
            Do Until .EOF
                If LocalCobranca.LocalizarCodigo(!local_cobranca) Then
                    x_local_cobranca = LocalCobranca.Nome
                Else
                    x_local_cobranca = "***"
                End If
                If lDataSubTotal <> !Data_Vencimento Then
                    ImpSubTotal
                    lDataSubTotal = !Data_Vencimento
                    lSubTotal = 0
                End If
                lSubTotal = lSubTotal + !Valor
                l_total = l_total + !Valor
                Call ImpDet(!Empresa, !Numero_Documento, x_local_cobranca, !Data_Emissao, !Data_Vencimento, !Valor, !Nome_Fornecedor, !Complemento)
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub PreencheCboConta()
    Dim rstConta As New adodb.Recordset
    
    cbo_conta.Clear
    cbo_conta.AddItem "Todas as Contas"
    cbo_conta.ItemData(cbo_conta.NewIndex) = 0
    Set rstConta = Conectar.RsConexao("SELECT Codigo, Nome FROM Contas ORDER BY Nome")
    'loop RecordSet
    With rstConta
        If Not .BOF Or Not .EOF Then
            .MoveFirst
            Do Until .EOF
                cbo_conta.AddItem !Nome
                cbo_conta.ItemData(cbo_conta.NewIndex) = !Codigo
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rstConta = Nothing
End Sub
Private Sub PreencheCboFornecedor()
    Dim rstFornecedor As New adodb.Recordset
    
    cbo_fornecedor.Clear
    cbo_fornecedor.AddItem "Todos os Fornecedores"
    cbo_fornecedor.ItemData(cbo_fornecedor.NewIndex) = 0
    Set rstFornecedor = Conectar.RsConexao("SELECT Codigo, Nome FROM Fornecedor WHERE Empresa = " & g_empresa & "ORDER BY Nome")
    'loop RecordSet
    With rstFornecedor
        If Not .BOF Or Not .EOF Then
            .MoveFirst
            Do Until .EOF
                cbo_fornecedor.AddItem !Nome
                cbo_fornecedor.ItemData(cbo_fornecedor.NewIndex) = !Codigo
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rstFornecedor = Nothing
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    Call LoopMovimentoContasPagar
    If l_total > 0 Then
        CalculaTotalEmpresas
        ImpSubTotal
        ImpTotal
        If lTotal(17, 11) > 0 Then
            Call ImpTotalEmpresas
        End If
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Contas à Pagar Especial|@|"
        frm_preview.Show 1
    Else
        MsgBox "Não existe movimento no período informado!", vbInformation, "Mensagem do Sistema"
    End If
    cmd_sair.SetFocus
End Sub
Private Sub ImpDet(ByVal pEmpresa As Integer, ByVal pNumeroDocumento As String, ByVal pLocalCobranca As String, ByVal pDataEmissao As Date, ByVal pDataVencimento As Date, ByVal pValor As Currency, ByVal pFornecedor As String, ByVal pComplemento As String)
    Dim x_linha As String
    Dim i As Integer
    
    If lPagina = 0 Then
        ImpCab
        ImpCab2
    End If
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    If lLinha >= 60 Then
        x_linha = "+------------+------------+----------+----------+---------------+------------------------------------------+-------------------------+--+"
        Mid(x_linha, 73, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.NewPage"
        ImpCab
        ImpCab2
    End If
    '                   1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '          12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    x_linha = "|            |            |          |          |               |                                          |                         |  |"
    Mid(x_linha, 3, 10) = pNumeroDocumento
    Mid(x_linha, 16, 10) = pLocalCobranca
    Mid(x_linha, 28, 10) = pDataEmissao
    Mid(x_linha, 39, 10) = pDataVencimento
    i = Len(Format(pValor, "######,##0.00"))
    Mid(x_linha, 51 + 13 - i, i) = Format(pValor, "######,##0.00")
    Mid(x_linha, 67, 40) = pFornecedor
    Mid(x_linha, 110, 23) = pComplemento
    i = Len(Format(pEmpresa, "#0"))
    Mid(x_linha, 135 + 2 - i, i) = Format(pEmpresa, "#0")
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    lLinha = lLinha + 1
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
    x_linha = "|                                                                  Página,     |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    Mid(x_linha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| CONTAS À PAGAR (ESPECIAL)                                       , __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(x_linha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(x_linha, 69, 10) = msk_data
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| PERÍODO DE VENCIMENTO...: __/__/____ A __/__/____                            |"
    Mid(x_linha, 29, 10) = msk_data_i
    Mid(x_linha, 42, 10) = msk_data_f
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| CONTA...................:                                                    |"
    Mid(x_linha, 29, 40) = cbo_conta
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| FORNECEDOR..............:                                                    |"
    Mid(x_linha, 29, 40) = cbo_fornecedor
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
End Sub
Private Sub ImpCab2()
    BioImprime "@Printer.Print " & "+------------+------------+----------+----------+---------------+------------------------------------------+-------------------------+--+"
    BioImprime "@Printer.Print " & "| NUMERO  DO | LOCAL   DE | DATA  DA | DATA  DO |   VALOR   DO  | NOME DO FORNECEDOR                       | COMPLEMENTO DO          |EM|"
    BioImprime "@Printer.Print " & "|  DOCUMENTO |  COBRANCA  |  EMISSAO |VENCIMENTO|   VENCIMENTO  |                                          | HISTORICO               |  |"
    BioImprime "@Printer.Print " & "+------------+------------+----------+----------+---------------+------------------------------------------+-------------------------+--+"
End Sub
Private Sub ImpSubTotal()
    Dim x_linha As String
    Dim i As Integer
    
    If lSubTotal > 0 Then
        If lLinha >= 60 Then
            x_linha = "+------------+------------+----------+----------+---------------+------------------------------------------+-------------------------+--+"
            Mid(x_linha, 73, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & x_linha
            BioImprime "@@Printer.NewPage"
            ImpCab
            ImpCab2
        End If
        x_linha = "|                  *** TOTAL DA DATA |--/--/----|               |                                                                    |  |"
        Mid(x_linha, 39, 10) = Format(lDataSubTotal, "dd/mm/yyyy")
        i = Len(Format(lSubTotal, "######,##0.00"))
        Mid(x_linha, 51 + 13 - i, i) = Format(lSubTotal, "######,##0.00")
        BioImprime "@@Printer.FontBold = True"
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.FontBold = False"
        BioImprime "@Printer.Print " & "+------------+------------+----------+----------+---------------+------------------------------------------+-------------------------+--+"
        lLinha = lLinha + 2
    End If
End Sub
Private Sub ImpTotal()
    Dim x_linha As String
    Dim i As Integer
    'BioImprime "@Printer.Print " & "+------------+------------+----------+----------+---------------+------------------------------------------+-------------------------+--+"
    x_linha = "|                *** TOTAL GERAL DO RELATORIO   |               |                                                                       |"
    i = Len(Format(l_total, "######,##0.00"))
    Mid(x_linha, 51 + 13 - i, i) = Format(l_total, "######,##0.00")
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    If lTotal(17, 11) = 0 Then
        x_linha = "+-----------------------------------------------+---------------+-----------------------------------------------------------------------+"
        BioImprime "@Printer.Print " & x_linha
    End If
End Sub
Private Sub ImpTotalEmpresas()
    Dim x_linha As String
    Dim x_empresa As Integer
    Dim i As Integer
    Dim i2 As Integer
    Dim i3 As Currency
    '                                        1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '                               12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    If lLinha >= 26 Then
        x_linha = "+-----------------------------------------------+---------------+-----------------------------------------------------------------------+"
        Mid(x_linha, 73, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.NewPage"
        ImpCab
        BioImprime "@Printer.Print " & "+------------+------------+------------+------------+------------+------------+------------+------------+------------+------------------+"
    Else
        BioImprime "@Printer.Print " & "+------------+------------+------------+--------+---+-----------++------------+------------+------------+------------+------------------+"
    End If
    x_linha = "| VENCIMENTO |            |            |            |            |            |            |            |            |                  |"
    For x_empresa = 1 To 8
        i2 = Len(Trim(lNomeEmpresa(x_empresa)))
        i3 = (12 * x_empresa + x_empresa + 2) + ((12 - i2) / 2)
        If Mid(Format(i3, "000.0"), 5, 1) <> "0" Then
            i = Val(i3) + 1
        Else
            i = Val(i3)
        End If
        Mid(x_linha, i, i2) = Trim(lNomeEmpresa(x_empresa))
    Next
    BioImprime "@Printer.Print " & x_linha
    
    x_linha = "|            |            |            |            |            |            |            |            |            |     T O T A L    |"
    For x_empresa = 9 To 16
        i2 = Len(Trim(lNomeEmpresa(x_empresa)))
        i3 = (12 * (x_empresa - 8) + (x_empresa - 8) + 2) + ((12 - i2) / 2)
        If Mid(Format(i3, "000.0"), 5, 1) <> "0" Then
            i = Val(i3) + 1
        Else
            i = Val(i3)
        End If
        Mid(x_linha, i, i2) = Trim(lNomeEmpresa(x_empresa))
    Next
    BioImprime "@Printer.Print " & x_linha
    
    
    BioImprime "@Printer.Print " & "+------------+------------+------------+------------+------------+------------+------------+------------+------------+------------------+"
    BioImprime "@@Printer.FontBold = True"
    For i = 1 To 10
        x_linha = "|            |            |            |            |            |            |            |            |            |                  |"
        For x_empresa = 1 To 8
            Mid(x_linha, 3, 10) = lData(i)
            i2 = Len(Format(lTotal(x_empresa, i), "#####,##0.00"))
            Mid(x_linha, (12 * x_empresa + x_empresa + 2) + 12 - i2, i2) = Format(lTotal(x_empresa, i), "#####,##0.00")
        Next
        'i2 = Len(Format(lTotal(x_empresa, i), "#####,##0.00"))
        'Mid(x_linha, 122 + 12 - i2, i2) = Format(lTotal(x_empresa, i), "#####,##0.00")
        BioImprime "@Printer.Print " & x_linha
    
    
        x_linha = "|            |            |            |            |            |            |            |            |            |                  |"
        For x_empresa = 9 To 16
            Mid(x_linha, 3, 10) = lData(i)
            i2 = Len(Format(lTotal(x_empresa, i), "#####,##0.00"))
            Mid(x_linha, (12 * (x_empresa - 8) + (x_empresa - 8) + 2) + 12 - i2, i2) = Format(lTotal(x_empresa, i), "#####,##0.00")
        Next
        i2 = Len(Format(lTotal(x_empresa, i), "#####,##0.00"))
        Mid(x_linha, 122 + 12 - i2, i2) = Format(lTotal(x_empresa, i), "#####,##0.00")
        BioImprime "@Printer.Print " & x_linha
        x_linha = "+------------+------------+------------+------------+------------+------------+------------+------------+------------+------------------+"
        BioImprime "@Printer.Print " & x_linha
    
    
    Next
    BioImprime "@@Printer.FontBold = False"
    'x_linha = "+------------+------------+------------+------------+------------+------------+------------+------------+------------+------------------+"
    'BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|  *** TOTAL |            |            |            |            |            |            |            |            |                  |"
    For x_empresa = 1 To 8
        i2 = Len(Format(lTotal(x_empresa, 11), "#####,##0.00"))
        Mid(x_linha, (12 * x_empresa + x_empresa + 2) + 12 - i2, i2) = Format(lTotal(x_empresa, 11), "#####,##0.00")
    Next
    'i2 = Len(Format(lTotal(x_empresa, 11), "#####,##0.00"))
    'Mid(x_linha, 122 + 12 - i2, i2) = Format(lTotal(x_empresa, 11), "#####,##0.00")
    BioImprime "@Printer.Print " & x_linha
    
    x_linha = "|            |            |            |            |            |            |            |            |            |                  |"
    For x_empresa = 9 To 16
        i2 = Len(Format(lTotal(x_empresa, 11), "#####,##0.00"))
        Mid(x_linha, (12 * (x_empresa - 8) + (x_empresa - 8) + 2) + 12 - i2, i2) = Format(lTotal(x_empresa, 11), "#####,##0.00")
    Next
    i2 = Len(Format(lTotal(x_empresa, 11), "#####,##0.00"))
    Mid(x_linha, 122 + 12 - i2, i2) = Format(lTotal(x_empresa, 11), "#####,##0.00")
    BioImprime "@Printer.Print " & x_linha
    
    BioImprime "@@Printer.FontBold = False"
    x_linha = "+------------+------------+------------+------------+------------+------------+------------+------------+------------+------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ZeraVariaveis()
    Dim i As Integer
    Dim i2 As Integer
    Dim rstEmpresa As New adodb.Recordset
    
    lLinha = 0
    lPagina = 0
    l_total = 0
    lDataSubTotal = CDate("00:00:00")
    lSubTotal = 0
    For i = 1 To 17
        For i2 = 1 To 11
            lTotal(i, i2) = 0
        Next
    Next
    For i = 1 To 16
        lCodigoEmpresa(i) = 15
        lNomeEmpresa(i) = ""
    Next
    g_string = ReadINI("GRUPO DE EMPRESAS", "Nome Abreviado das Empresas", gArquivoIni)
    Set rstEmpresa = Conectar.RsConexao("SELECT Codigo, Nome FROM Empresas WHERE Inativo = " & preparaBooleano(False) & "ORDER BY Codigo")
    With rstEmpresa
        If .RecordCount > 0 Then
            .MoveFirst
            i = 0
            Do Until .EOF
                i = i + 1
                If i > 16 Then
                    Exit Do
                End If
                lCodigoEmpresa(i) = !Codigo
                lNomeEmpresa(i) = RetiraGString(i)
                If GrupoEmpresa.LocalizarGrupoEmpresa(1, !Codigo) Then
                    lNomeEmpresa(i) = GrupoEmpresa.NomeEmpresaReduzido
                End If
                .MoveNext
            Loop
        End If
    End With
    rstEmpresa.Close
    Set rstEmpresa = Nothing
    g_string = ""
    For i = 1 To 10
        lData(i) = CDate(msk_data.Text) - 1 + i
    Next
End Sub
Private Sub cbo_conta_GotFocus()
    SendMessageLong cbo_conta.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_conta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_fornecedor.SetFocus
    End If
End Sub
Private Sub cbo_fornecedor_GotFocus()
    SendMessageLong cbo_fornecedor.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_fornecedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub

Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cbo_conta.SetFocus
    Else
        msk_data = RetiraGString(1)
        msk_data_i.SetFocus
    End If
    g_string = ""
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
    g_string = ""
    cbo_conta.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cbo_conta.SetFocus
    Else
        msk_data_i = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_imprimir_Click()
    lLocal = 1
    If ValidaCampos Then
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
            rstMovimentoContasPagar.Close
            Set rstMovimentoContasPagar = Nothing
        End If
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    Dim x_data_i As Date
    Dim x_data_f As Date
    x_data_i = msk_data_i
    x_data_f = msk_data_f
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf Not IsDate(msk_data_i) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf x_data_f < x_data_i Then
        MsgBox "Data final deve ser maior que a data inicial.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf cbo_conta.ListIndex = -1 Then
        MsgBox "Selecione uma conta.", vbInformation, "Atenção!"
        cbo_conta.SetFocus
    ElseIf cbo_fornecedor.ListIndex = -1 Then
        MsgBox "Selecione um fornecedor.", vbInformation, "Atenção!"
        cbo_fornecedor.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Finaliza
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
Private Sub cmd_visualizar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_imprimir_Click
    ElseIf KeyCode = vbKeyF9 Then
        KeyCode = 0
        cmd_visualizar_Click
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If g_string = "" Then
        If Not IsDate(msk_data.Text) Then
            msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
            msk_data_i.Text = Format(MovimentoContaPagar.PrimeiraDataVencimento(g_empresa), "dd/mm/yyyy")  '  Format(g_data_def, "dd/mm/yyyy")
            msk_data_f.Text = Format(g_data_def + 90, "dd/mm/yyyy")
            cbo_conta.ListIndex = 0
            cbo_fornecedor.ListIndex = 0
        End If
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    
    PreencheCboConta
    PreencheCboFornecedor
End Sub
Private Sub msk_data_f_GotFocus()
    msk_data_f.SelStart = 0
    msk_data_f.SelLength = 5
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_conta.SetFocus
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
