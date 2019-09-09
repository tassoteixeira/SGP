VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_nota_cliente_geral 
   Caption         =   "Emissão das Notas de Abastecimento/Duplicata"
   ClientHeight    =   3405
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_nota_cliente_geral.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_nota_cliente_geral.frx":030A
   ScaleHeight     =   3405
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_nota_cliente_geral.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Visualiza notas de abastecimento por emissão."
      Top             =   2460
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_nota_cliente_geral.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Imprime notas de abastecimento por emissão."
      Top             =   2460
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_nota_cliente_geral.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2460
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CheckBox chkConsiderarNotasBaixadas 
         Caption         =   "Considerar Notas Baixadas"
         Height          =   255
         Left            =   3600
         TabIndex        =   18
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CheckBox chkSomenteCombustivel 
         Caption         =   "Somente notas de combustível"
         Height          =   195
         Left            =   3600
         TabIndex        =   17
         Top             =   1560
         Width           =   2775
      End
      Begin VB.ComboBox cboGrupoCliente 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1800
         Width           =   4755
      End
      Begin VB.CheckBox chkUnificaEmpresa 
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   1440
         Width           =   435
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_nota_cliente_geral.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_nota_cliente_geral.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_nota_cliente_geral.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   1080
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
         Caption         =   "&Grupo de Cliente"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Unifica empresas"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1440
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
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   2700
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_nota_cliente_geral"
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
Dim lSQL As String
'Fim de variáveis padrão para relatório
Dim lTotal As Currency
Dim lTotalLitroCombustivel As Currency
Dim lValorAbastecimento As Currency
Dim lLitrosAbastecimento As Currency
Dim lTotalDup As Currency
Dim lTotalDupDia As Currency
Dim lTotalDupVencer As Currency
Dim lTotalDupVencida As Currency
Dim lTotalCustoBancario As String
Dim lQtdDup As Currency
Dim lQtdDupDia As Currency
Dim lQtdDupVencer As Currency
Dim lQtdDupVencida As Currency

Private Cliente As New cCliente
Private DuplicataReceber As New cDuplicataReceber
Private MovimentoNotaAbastecimento As New cMovimentoNotaAbastecimento
Private rsTabela As New ADODB.Recordset
Private rsTabelaNotaAbastecimento As New ADODB.Recordset
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Cliente = Nothing
    Set DuplicataReceber = Nothing
    Set MovimentoNotaAbastecimento = Nothing
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lTotal = 0
    lTotalLitroCombustivel = 0
    lTotalDup = 0
    lTotalDupDia = 0
    lTotalDupVencer = 0
    lTotalDupVencida = 0
    lTotalCustoBancario = 0
    lQtdDup = 0
    lQtdDupDia = 0
    lQtdDupVencer = 0
    lQtdDupVencida = 0
End Sub
Private Sub AtivaBotoes(ByVal pAtiva As Boolean)
    cmd_visualizar.Enabled = pAtiva
    cmd_imprimir.Enabled = pAtiva
    cmd_sair.Enabled = pAtiva
    If pAtiva = False Then
        frmAguarde.Show
        Call frmAguarde.MostraMensagens("Gerando Relatório!", Me.Caption, Me.Top, Me.Left, Me.Width, Me.Height)
        DoEvents
    Else
        Call frmAguarde.Finaliza
    End If
End Sub
Private Sub BuscaDatas()
    Dim xDataI As Date
    Dim xDataF As Date
    Dim xData As Date
    Dim xEmpresa As Integer
    
    xDataI = Date
    xDataF = Date
    
    xEmpresa = g_empresa
    If chkUnificaEmpresa.Value = 1 Then
        xEmpresa = 0
    End If
    
    
    'Busca datas das Duplicatas
    xData = DuplicataReceber.LocalizaPrimeiroVencimento(xEmpresa)
    If xData < xDataI Then
        xDataI = xData
    End If
    xData = DuplicataReceber.LocalizaUltimoVencimento(xEmpresa)
    If xData > xDataF Then
        xDataF = xData
    End If
        
    
    'Busca datas das Notas de Abastecimento
    xData = MovimentoNotaAbastecimento.LocalizaPrimeiraData(xEmpresa)
    If xData < xDataI Then
        xDataI = xData
    End If
    xData = MovimentoNotaAbastecimento.LocalizaUltimaData(xEmpresa)
    If xData > xDataF Then
        xDataF = xData
    End If

    msk_data_i.Text = Format(xDataI, "dd/mm/yyyy")
    msk_data_f.Text = Format(xDataF, "dd/mm/yyyy")
End Sub
Private Sub PreencheCboGrupoCliente()
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Codigo, Nome"
    lSQL = lSQL & "     FROM GrupoCliente"
    lSQL = lSQL & " ORDER BY Nome, Codigo"
    'Abre RecordSet
    Set rsTabela = New ADODB.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    
    cboGrupoCliente.Clear
    cboGrupoCliente.AddItem "Todos os Grupos"
    cboGrupoCliente.ItemData(cboGrupoCliente.NewIndex) = 0
    If rsTabela.RecordCount > 0 Then
        rsTabela.MoveFirst
        Do Until rsTabela.EOF
            cboGrupoCliente.AddItem rsTabela("Nome").Value
            cboGrupoCliente.ItemData(cboGrupoCliente.NewIndex) = rsTabela("Codigo").Value
            rsTabela.MoveNext
        Loop
    End If
    rsTabela.Close
    Set rsTabela = Nothing
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Codigo, [Razao Social] as NomeCliente "
    lSQL = lSQL & "  FROM Cliente"
    If cboGrupoCliente.ItemData(cboGrupoCliente.ListIndex) > 0 Then
        lSQL = lSQL & "    WHERE [Codigo do Grupo de Cliente] = " & cboGrupoCliente.ItemData(cboGrupoCliente.ListIndex)
    End If
    lSQL = lSQL & " ORDER BY [Razao Social]"
    'Abre RecordSet
    Set rsTabela = New ADODB.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
    
        If CBool(chkSomenteCombustivel.Value) = True Then
            Call ImpDadosCombustivel(chkConsiderarNotasBaixadas.Value)
        Else
            ImpDados
        End If
    
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
    RelatorioDuplicata
    If lPagina > 0 Then
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Notas de Abastecimento/Duplicata|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub RelatorioDuplicata()
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Duplicata_Receber.[Numero do Documento], Duplicata_Receber.[Codigo do Cliente],"
    lSQL = lSQL & "       Duplicata_Receber.[Data do Vencimento], Duplicata_Receber.[Valor do Vencimento],"
    lSQL = lSQL & "       Duplicata_Receber.[Valor do Custo Bancario], Cliente.[Razao Social],"
    lSQL = lSQL & "       Cliente.Telefone"
    lSQL = lSQL & "  FROM Duplicata_Receber, Cliente"
    lSQL = lSQL & " WHERE Duplicata_Receber.[Data do Vencimento] >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "   AND Duplicata_Receber.[Data do Vencimento] <= " & preparaData(CDate(msk_data_f.Text))
    If Me.chkUnificaEmpresa.Value = 0 Then
        lSQL = lSQL & "   AND Duplicata_Receber.Empresa = " & g_empresa
    End If
    lSQL = lSQL & "   AND Duplicata_Receber.[Codigo do Cliente] = Cliente.Codigo"
    lSQL = lSQL & " ORDER BY Cliente.[Razao Social] ASC, Duplicata_Receber.[Data do Vencimento] ASC"
    'Abre RecordSet
    Set rsTabela = New ADODB.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        ImpDadosDuplicata
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
End Sub
Private Sub ImpDados()
    Dim xLinha As String
    Dim xEmpresa As Integer
    
    If chkUnificaEmpresa.Value = 1 Then
        xEmpresa = 0
    Else
        xEmpresa = g_empresa
    End If
    Do Until rsTabela.EOF
        If lPagina = 0 Then
            ImpCab
            ImpCabNota
        End If
        If lLinha >= 55 Then
            xLinha = "+-------+------------------------------------------+----------------+----------+"
            Mid(xLinha, 25, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & xLinha
            BioImprime "@@Printer.NewPage"
            ImpCab
            ImpCabNota
        End If
        lValorAbastecimento = MovimentoNotaAbastecimento.TotalDataLiquido(xEmpresa, rsTabela("Codigo").Value, CDate(msk_data_i.Text), CDate(msk_data_f.Text), False, CBool(chkConsiderarNotasBaixadas.Value))
        If lValorAbastecimento > 0 Then
            ImpCliente
            lTotal = lTotal + lValorAbastecimento
        End If
        rsTabela.MoveNext
    Loop
    If lTotal > 0 Then
        ImpTotal
    End If
End Sub
Private Sub ImpDadosCombustivel(ByVal pConsiderarBaixadas As Boolean)
    Dim xLinha As String
    Dim xEmpresa As Integer
    
    If chkUnificaEmpresa.Value = 1 Then
        xEmpresa = 0
    Else
        xEmpresa = g_empresa
    End If
    Do Until rsTabela.EOF
        If lPagina = 0 Then
            ImpCab
            ImpCabNotaCombustivel
        End If
        If lLinha >= 55 Then
            xLinha = "+-------+------------------------------------------+-----------+-----------+----------+"
            Mid(xLinha, 25, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & xLinha
            BioImprime "@@Printer.NewPage"
            ImpCab
            ImpCabNotaCombustivel
        End If
        lValorAbastecimento = 0 'MovimentoNotaAbastecimento.TotalDataLiquido(xEmpresa, rsTabela("Codigo").Value, CDate(msk_data_i.Text), CDate(msk_data_f.Text), False)
        lLitrosAbastecimento = 0 'MovimentoNotaAbastecimento.TotalDataLitrosCombustivel(xEmpresa, rsTabela("Codigo").Value, CDate(msk_data_i.Text), CDate(msk_data_f.Text), False)
        
        Call BuscaTotaisLitroEValorLiquidoCombustivelCliente(xEmpresa, rsTabela("Codigo").Value, CDate(msk_data_i.Text), CDate(msk_data_f.Text), chkConsiderarNotasBaixadas.Value)
        
        If lValorAbastecimento > 0 Then
            ImpClienteCombustivel
            lTotal = lTotal + lValorAbastecimento
            lTotalLitroCombustivel = lTotalLitroCombustivel + lLitrosAbastecimento
        End If
        rsTabela.MoveNext
    Loop
    If lTotal > 0 Then
        ImpTotalCombustivel
    End If
End Sub

Private Sub BuscaTotaisLitroEValorLiquidoCombustivelCliente(ByVal pEmpresa As Integer, ByVal pCodigoCliente As Long, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pCondiderarBaixadas As Boolean)

    Dim xSQL As String

    Dim xValor As Currency
    Dim xQuantidade As Currency
    
    Set rsTabelaNotaAbastecimento = New ADODB.Recordset


On Error GoTo trata_erro

    xValor = 0
    xQuantidade = 0
    xSQL = "SELECT Quantidade, [Valor Desconto Unitario], [Valor Unitario], [Valor Total] FROM Movimento_Nota_Abastecimento"
    If pEmpresa > 0 Then
        xSQL = xSQL & " WHERE Empresa = " & pEmpresa
    Else
        xSQL = xSQL & " WHERE Empresa > 0"
    End If
    If pCodigoCliente > 0 Then
        xSQL = xSQL & " AND [Codigo do Cliente] = " & pCodigoCliente
    End If
    
    'Forma de selecionar somente nota de combustivel
    xSQL = xSQL & " AND [Codigo do Produto2] IN (SELECT DISTINCT([Codigo do Produto]) FROM Bomba WHERE Empresa =" & pEmpresa & ")"
        
    xSQL = xSQL & " AND [Data do Abastecimento] >= " & preparaData(pDataInicial)
    xSQL = xSQL & " AND [Data do Abastecimento] <= " & preparaData(pDataFinal)
    
    If pCondiderarBaixadas Then
    
        xSQL = xSQL & " UNION "
        xSQL = xSQL & " SELECT Quantidade, [Valor Desconto Unitario], [Valor Unitario], [Valor Total] FROM Baixa_Nota_Abastecimento"
        If pEmpresa > 0 Then
            xSQL = xSQL & " WHERE Empresa = " & pEmpresa
        Else
            xSQL = xSQL & " WHERE Empresa > 0"
        End If
        If pCodigoCliente > 0 Then
            xSQL = xSQL & " AND [Codigo do Cliente] = " & pCodigoCliente
        End If
            
        'Forma de selecionar somente nota de combustivel
        xSQL = xSQL & " AND [Codigo do Produto2] IN (SELECT DISTINCT([Codigo do Produto]) FROM Bomba WHERE Empresa =" & pEmpresa & ")"
    
            
        xSQL = xSQL & " AND [Data do Abastecimento] >= " & preparaData(pDataInicial)
        xSQL = xSQL & " AND [Data do Abastecimento] <= " & preparaData(pDataFinal)
    
    End If
    
    Set rsTabelaNotaAbastecimento = Conectar.RsConexao(xSQL)

    
    If rsTabelaNotaAbastecimento.RecordCount > 0 Then
            rsTabelaNotaAbastecimento.MoveFirst
            Do Until rsTabelaNotaAbastecimento.EOF
                'Quantidade
                xQuantidade = xQuantidade + rsTabelaNotaAbastecimento("Quantidade").Value
                'Desconto
                If rsTabelaNotaAbastecimento("Valor Desconto Unitario").Value > 0 Then
                    xValor = xValor + rsTabelaNotaAbastecimento("Valor Total").Value
                    xValor = xValor - Format(rsTabelaNotaAbastecimento("Valor Desconto Unitario").Value * rsTabelaNotaAbastecimento("Quantidade").Value, "0000000000.00")
                'Acrescimo
                ElseIf rsTabelaNotaAbastecimento("Valor Desconto Unitario").Value < 0 Then
                    xValor = xValor + rsTabelaNotaAbastecimento("Valor Total").Value
                    xValor = xValor + Format(rsTabelaNotaAbastecimento("Valor Desconto Unitario").Value * -1 * rsTabelaNotaAbastecimento("Quantidade").Value, "0000000000.00")
                Else
                    xValor = xValor + rsTabelaNotaAbastecimento("Valor Total").Value
                End If
                rsTabelaNotaAbastecimento.MoveNext
            Loop
    End If
    rsTabelaNotaAbastecimento.Close
    Set rsTabelaNotaAbastecimento = Nothing
    
    lValorAbastecimento = xValor
    lLitrosAbastecimento = xQuantidade

    Exit Sub

trata_erro:
    MsgBox Err.Number & " - " & Err.Description



End Sub

Private Sub ImpDadosDuplicata()
    Dim xLinha As String
    
    If lPagina = 0 Then
        ImpCab
    End If
    ImpCabDuplicata
    Do Until rsTabela.EOF
        If lLinha >= 55 Then
            xLinha = "+-------+------------------------------------------+----------------+----------+"
            Mid(xLinha, 25, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & xLinha
            BioImprime "@@Printer.NewPage"
            ImpCab
            ImpCabDuplicata
        End If
        ImpDetDuplicata
        lQtdDup = lQtdDup + 1
        lTotalDup = lTotalDup + rsTabela("Valor do Vencimento").Value
        lTotalCustoBancario = lTotalCustoBancario + rsTabela("Valor do Custo Bancario").Value
        If rsTabela("Data do Vencimento").Value = CDate(msk_data.Text) Then
            lTotalDupDia = lTotalDupDia + rsTabela("Valor do Vencimento").Value
            lQtdDupDia = lQtdDupDia + 1
        ElseIf rsTabela("Data do Vencimento").Value > CDate(msk_data.Text) Then
            lTotalDupVencer = lTotalDupVencer + rsTabela("Valor do Vencimento").Value
            lQtdDupVencer = lQtdDupVencer + 1
        Else
            lTotalDupVencida = lTotalDupVencida + rsTabela("Valor do Vencimento").Value
            lQtdDupVencida = lQtdDupVencida + 1
        End If
        rsTabela.MoveNext
    Loop
    If lTotalDup > 0 Then
        ImpTotalDuplicata
    End If
End Sub
Private Sub ImpCliente()
    Dim xLinha As String
    Dim i As Integer
    Dim xData As Date
    Dim xDias As Integer
    
    xLinha = "|       |                                          |                |          |"
    i = Len(Format(rsTabela("Codigo").Value, "#####"))
    Mid(xLinha, 3 + 5 - i, i) = Format(rsTabela("Codigo").Value, "#####")
    Mid(xLinha, 11, 40) = rsTabela("NomeCliente").Value
    i = Len(Format(lValorAbastecimento, "###,###,##0.00"))
    Mid(xLinha, 54 + 14 - i, i) = Format(lValorAbastecimento, "###,###,##0.00")
    
    xData = MovimentoNotaAbastecimento.LocalizaPrimeiraDataCliente(g_empresa, CLng(rsTabela("Codigo").Value))
    xDias = DateDiff("d", xData, Date)
    i = Len(Format(xDias, "##,##0"))
    Mid(xLinha, 72 + 6 - i, i) = Format(xDias, "##,##0")
    
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpClienteCombustivel()
    Dim xLinha As String
    Dim i As Integer
    Dim xData As Date
    Dim xDias As Integer
    '        "+----+------------------------------------------+------------+--------------+---------+"
    xLinha = "|    |                                          |            |              |         |"
    i = Len(Format(rsTabela("Codigo").Value, "#####"))
    Mid(xLinha, 2 + 4 - i, i) = Format(rsTabela("Codigo").Value, "#####")
    Mid(xLinha, 7, 40) = rsTabela("NomeCliente").Value
    
    i = Len(Format(lLitrosAbastecimento, "###,###,##0.00"))
    Mid(xLinha, 50 + 12 - i, i) = Format(lLitrosAbastecimento, "###,###,##0.00")
    
    i = Len(Format(lValorAbastecimento, "###,###,##0.00"))
    Mid(xLinha, 63 + 14 - i, i) = Format(lValorAbastecimento, "###,###,##0.00")
    
    xData = MovimentoNotaAbastecimento.LocalizaPrimeiraDataCliente(g_empresa, CLng(rsTabela("Codigo").Value))
    xDias = DateDiff("d", xData, Date)
    i = Len(Format(xDias, "##,##0"))
    Mid(xLinha, 78 + 9 - i, i) = Format(xDias, "##,##0")
    
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpDetDuplicata()
    Dim xLinha As String
    Dim i As Integer
    Dim xString As String
    
    '                                         1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '                                12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    'BioImprime "@Printer.Print " & "+----------+----+----------------------------------------+----------+-------+------------+----------------------------+-----------------+"
    'BioImprime "@Printer.Print " & "| NUMERO DO|CÓD.|RAZÃO SOCIAL                            |DATA    DO| CUSTO |  VALOR DO  |          SITUACAO          | TELEFONE        |"
    'BioImprime "@Printer.Print " & "| DOCUMENTO|CLI.|                                        |VENCIMENTO|BANCAR.| VENCIMENTO |                            |                 |"
    'BioImprime "@Printer.Print " & "+----------+----+----------------------------------------+----------+-------+------------+----------------------------+-----------------+"
    xLinha = "|          |    |                                        |          |       |            |                            |                 |"
    i = Len(Format(rsTabela("Numero do Documento").Value, "#,###,###"))
    Mid(xLinha, 2 + 9 - i, i) = Format(rsTabela("Numero do Documento").Value, "#,###,###")
    i = Len(Format(rsTabela("Codigo do Cliente").Value, "####"))
    Mid(xLinha, 13 + 4 - i, i) = Format(rsTabela("Codigo do Cliente").Value, "####")
    Mid(xLinha, 18, 40) = rsTabela("Razao Social").Value
    Mid(xLinha, 59, 10) = Format(rsTabela("Data do Vencimento").Value, "dd/mm/yyyy")
    If rsTabela("Valor do Custo Bancario").Value > 0 Then
        i = Len(Format(rsTabela("Valor do Custo Bancario").Value, "###0.00"))
        Mid(xLinha, 70 + 7 - i, i) = Format(rsTabela("Valor do Custo Bancario").Value, "###0.00")
    End If
    i = Len(Format(rsTabela("Valor do Vencimento").Value, "#,###,##0.00"))
    Mid(xLinha, 78 + 12 - i, i) = Format(rsTabela("Valor do Vencimento").Value, "#,###,##0.00")
    If rsTabela("Data do Vencimento").Value = CDate(msk_data.Text) Then
        xString = "Vence Hoje"
    ElseIf rsTabela("Data do Vencimento").Value > CDate(msk_data.Text) Then
        xString = "Vence em (" & Format(DateDiff("d", CDate(msk_data.Text), rsTabela("Data do Vencimento").Value), "000") & ") Dia(s)"
    Else
        xString = "Vencida à (" & Format(DateDiff("d", rsTabela("Data do Vencimento").Value, CDate(msk_data.Text)), "000") & ") Dia(s)"
    End If
    Mid(xLinha, 91, 28) = xString
    Mid(xLinha, 121, 16) = fMascaraTelefone(rsTabela("Telefone").Value)
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    xLinha = "+-------+------------------------------------------+----------------+----------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                                           TOTAL  |                |          |"
    i = Len(Format(lTotal, "###,###,##0.00"))
    Mid(xLinha, 54 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+--------------------------------------------------+----------------+----------+"
    Mid(xLinha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpTotalCombustivel()
    Dim xLinha As String
    Dim i As Integer
    
    '        "+----+------------------------------------------+------------+--------------+---------+"
    xLinha = "+----+------------------------------------------+------------+--------------+---------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                                        TOTAL  |            |              |         |"
    i = Len(Format(lTotalLitroCombustivel, "###,###,##0.00"))
    Mid(xLinha, 50 + 12 - i, i) = Format(lTotalLitroCombustivel, "###,###,##0.00")
    i = Len(Format(lTotal, "###,###,##0.00"))
    Mid(xLinha, 63 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-----------------------------------------------+------------+--------------+---------+"
    Mid(xLinha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpTotalDuplicata()
    Dim xLinha As String
    Dim i As Integer
    
    BioImprime "@Printer.Print " & "+----------+----+----------------------------------------+----------+-------+------------+----------------------------+-----------------+"
    
    xLinha = "|                                                                   |       |            |                                              |"
    i = Len(Format(lTotalDupVencida, "#,###,##0.00"))
    Mid(xLinha, 78 + 12 - i, i) = Format(lTotalDupVencida, "#,###,##0.00")
    Mid(xLinha, 91, 35) = " " & lQtdDupVencida & " DUPLICATA(S) VENCIDA(S)"
    BioImprime "@Printer.Print " & xLinha
    
    xLinha = "|                                                                   |       |            |                                              |"
    i = Len(Format(lTotalDupDia, "#,###,##0.00"))
    Mid(xLinha, 78 + 12 - i, i) = Format(lTotalDupDia, "#,###,##0.00")
    Mid(xLinha, 91, 35) = " " & lQtdDupDia & " DUPLICATA(S) VENCE(M) HOJE"
    BioImprime "@Printer.Print " & xLinha
    
    xLinha = "|                                                                   |       |            |                                              |"
    i = Len(Format(lTotalDupVencer, "#,###,##0.00"))
    Mid(xLinha, 78 + 12 - i, i) = Format(lTotalDupVencer, "#,###,##0.00")
    Mid(xLinha, 91, 35) = " " & lQtdDupVencer & " DUPLICATA(S) À VENCER"
    BioImprime "@Printer.Print " & xLinha
    
    xLinha = "|                                                                   |       |            |                                              |"
    i = Len(Format(lTotalCustoBancario, "###0.00"))
    Mid(xLinha, 70 + 7 - i, i) = Format(lTotalCustoBancario, "###0.00")
    i = Len(Format(lTotalDup, "#,###,##0.00"))
    Mid(xLinha, 78 + 12 - i, i) = Format(lTotalDup, "#,###,##0.00")
    Mid(xLinha, 91, 35) = " TOTAL GERAL DE " & lQtdDup & " DUPLICATAS"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-------------------------------------------------------------------+-------+------------+----------------------------------------------+"
    Mid(xLinha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpCab()
    Dim xLinha As String
    Dim i As Integer
    Dim x_string_40 As String * 40
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
    '                                   +-------+------------------------------------------+-----------+-----------+----------+
    If chkSomenteCombustivel.Value Then
        BioImprime "@Printer.Print " & "+-------------------------------------------------------------------------------------+"
        x_string_40 = g_nome_empresa
        BioImprime "@Printer.Print " & "| " & x_string_40 & "                                Página, " & Format(lPagina, "000") & " |"
        '                   1         2         3         4         5         6         7         8
        '          12345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                                              123456789012345678901234567890
        xLinha = "| NOTAS DE ABASTECIMENTO/DUPLICATA                                 CIDADE, __/__/____ |"
        i = Len(g_cidade_empresa)
        Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
        Mid(xLinha, 76, 10) = msk_data.Text
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@Printer.Print " & "| Referente a.: " & msk_data_i.Text & " a " & msk_data_f.Text & "                                               |"
        xLinha = "| GRUPO DE CLIENTE.:                                                                  |"
        Mid(xLinha, 22, 30) = cboGrupoCliente.Text
        BioImprime "@Printer.Print " & xLinha
        xLinha = "| Considerar Baixadas:                                       Somente Combustível:     |"
        Mid(xLinha, 24, 3) = IIf(CBool(chkConsiderarNotasBaixadas.Value), "SIM", "NÃO")
        Mid(xLinha, 83, 3) = IIf(CBool(chkSomenteCombustivel.Value), "SIM", "NÃO")
        BioImprime "@Printer.Print " & xLinha
    Else
        BioImprime "@Printer.Print " & "+------------------------------------------------------------------------------+"
        x_string_40 = g_nome_empresa
        BioImprime "@Printer.Print " & "| " & x_string_40 & "                         Página, " & Format(lPagina, "000") & " |"
        '                   1         2         3         4         5         6         7         8
        '          12345678901234567890123456789012345678901234567890123456789012345678901234567890
        '                                              123456789012345678901234567890
        xLinha = "| NOTAS DE ABASTECIMENTO/DUPLICATA                          CIDADE, __/__/____ |"
        i = Len(g_cidade_empresa)
        Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
        Mid(xLinha, 69, 10) = msk_data.Text
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@Printer.Print " & "| Referente a.: " & msk_data_i.Text & " a " & msk_data_f.Text & "                                        |"
        xLinha = "| GRUPO DE CLIENTE.:                                                           |"
        Mid(xLinha, 22, 30) = cboGrupoCliente.Text
        BioImprime "@Printer.Print " & xLinha
        xLinha = "| Considerar Baixadas:                                Somente Combustível:     |"
        Mid(xLinha, 24, 3) = IIf(CBool(chkConsiderarNotasBaixadas.Value), "SIM", "NÃO")
        Mid(xLinha, 76, 3) = IIf(CBool(chkSomenteCombustivel.Value), "SIM", "NÃO")
        BioImprime "@Printer.Print " & xLinha
    End If
    
End Sub
Private Sub ImpCabNota()
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+-------+------------------------------------------+----------------+----------+"
    BioImprime "@Printer.Print " & "|  COD. | RAZÃO SOCIAL                             | TOTAL   ABAST. |   DIAS   |"
    BioImprime "@Printer.Print " & "+-------+------------------------------------------+----------------+----------+"
    
    
End Sub
Private Sub ImpCabNotaCombustivel()
    BioImprime "@@Printer.FontBold = False"
    '                              "+----+------------------------------------------+------------+--------------+---------+"
    BioImprime "@Printer.Print " & "+----+------------------------------------------+------------+--------------+---------+"
    If CBool(chkConsiderarNotasBaixadas.Value) Then
        BioImprime "@Printer.Print " & "| COD| RAZÃO SOCIAL                             |     LITROS |       TOTAL  |DIAS S/BX|"
    Else
        BioImprime "@Printer.Print " & "| COD| RAZÃO SOCIAL                             |     LITROS |      TOTAL   |   DIAS  |"
    End If
    BioImprime "@Printer.Print " & "+----+------------------------------------------+------------+--------------+---------+"
    
    
End Sub
Private Sub ImpCabDuplicata()
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & "+----------+----+----------------------------------------+----------+-------+------------+----------------------------+-----------------+"
    BioImprime "@Printer.Print " & "| NUMERO DO|CÓD.|RAZÃO SOCIAL                            |DATA    DO| CUSTO |  VALOR DO  |          SITUACAO          | TELEFONE        |"
    BioImprime "@Printer.Print " & "| DOCUMENTO|CLI.|                                        |VENCIMENTO|BANCAR.| VENCIMENTO |                            |                 |"
    BioImprime "@Printer.Print " & "+----------+----+----------------------------------------+----------+-------+------------+----------------------------+-----------------+"
End Sub
Private Sub cboGrupoCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub chkUnificaEmpresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboGrupoCliente.SetFocus
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
    g_string = ""
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
    g_string = ""
End Sub
Private Sub cmd_imprimir_Click()
    lLocal = 1
    If ValidaCampos Then
        AtivaBotoes (False)
        If SelecionaImpressoraEpson(Me) Then
            DoEvents
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        End If
        AtivaBotoes (True)
        cmd_sair.SetFocus
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
        AtivaBotoes (False)
        If SelecionaImpressoraEpson(Me) Then
            DoEvents
            Call GravaAuditoria(1, Me.name, 6, "")
            Relatorio
        End If
        AtivaBotoes (True)
        cmd_sair.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        cboGrupoCliente.ListIndex = 0
        BuscaDatas
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
    PreencheCboGrupoCliente
End Sub
Private Sub msk_data_f_GotFocus()
    msk_data_f.SelStart = 0
    msk_data_f.SelLength = 2
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chkUnificaEmpresa.SetFocus
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
