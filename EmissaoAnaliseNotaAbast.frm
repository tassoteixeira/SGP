VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form EmissaoAnaliseNotaAbast 
   Caption         =   "Análise das Notas de Abastecimento"
   ClientHeight    =   2745
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "EmissaoAnaliseNotaAbast.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "EmissaoAnaliseNotaAbast.frx":030A
   ScaleHeight     =   2745
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "EmissaoAnaliseNotaAbast.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Visualiza notas de abastecimento por emissão."
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "EmissaoAnaliseNotaAbast.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprime notas de abastecimento por emissão."
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "EmissaoAnaliseNotaAbast.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1800
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "EmissaoAnaliseNotaAbast.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "EmissaoAnaliseNotaAbast.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2760
         Picture         =   "EmissaoAnaliseNotaAbast.frx":6CBA
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
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "EmissaoAnaliseNotaAbast"
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
Dim lTotalEntrada As Currency
Dim lTotalSaida As Currency
Dim lSaldoInicial As Currency
Dim lSaldoFinal As Currency

Private Cliente As New cCliente
Private DuplicataReceber As New cDuplicataReceber
Private MovimentoNotaAbastecimento As New cMovimentoNotaAbastecimento
Private rsTabela As New adodb.Recordset
Private rsTabela2 As New adodb.Recordset
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Cliente = Nothing
    Set DuplicataReceber = Nothing
    Set MovimentoNotaAbastecimento = Nothing
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lTotalEntrada = 0
    lTotalSaida = 0
    lSaldoInicial = 0
    lSaldoFinal = 0
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

    msk_data_i.Text = Format(Date, "dd/mm/yyyy")
    msk_data_f.Text = Format(Date, "dd/mm/yyyy")
End Sub
Private Sub CalculoInicial()
    'Calcula Entradas (Notas em Aberto no Período)
    lSQL = ""
    lSQL = lSQL & "SELECT SUM([Valor Total]) As Total"
    lSQL = lSQL & "  FROM Movimento_Nota_Abastecimento"
    lSQL = lSQL & " WHERE [Data do Abastecimento] >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND [Data do Abastecimento] <= " & preparaData(msk_data_f.Text)
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        If Not IsNull(rsTabela("Total").Value) Then
            lTotalEntrada = lTotalEntrada + rsTabela("Total").Value
        End If
    End If
    rsTabela.Close

    'Calcula Entradas (Notas Baixadas no Período)
    lSQL = ""
    lSQL = lSQL & "SELECT SUM([Valor Total]) As Total"
    lSQL = lSQL & "  FROM Baixa_Nota_Abastecimento"
    lSQL = lSQL & " WHERE [Data do Abastecimento] >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND [Data do Abastecimento] <= " & preparaData(msk_data_f.Text)
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        If Not IsNull(rsTabela("Total").Value) Then
            lTotalEntrada = lTotalEntrada + rsTabela("Total").Value
        End If
    End If
    rsTabela.Close

    'Calcula Saídas (Duplicatas Recebidas no Período
    lSQL = ""
    lSQL = lSQL & "SELECT SUM([Valor Pago]) As Total"
    lSQL = lSQL & "  FROM Baixa_Duplicata_Receber"
    lSQL = lSQL & " WHERE [Data do Pagamento] >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND [Data do Pagamento] <= " & preparaData(msk_data_f.Text)
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        If Not IsNull(rsTabela("Total").Value) Then
            lTotalSaida = lTotalSaida + rsTabela("Total").Value
        End If
    End If
    rsTabela.Close

    'Calcula Saldo Final (Total geral de Notas em Aberto)
    lSQL = ""
    lSQL = lSQL & "SELECT SUM([Valor Total]) As Total"
    lSQL = lSQL & "  FROM Movimento_Nota_Abastecimento"
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        If Not IsNull(rsTabela("Total").Value) Then
            lSaldoFinal = lSaldoFinal + rsTabela("Total").Value
        End If
    End If
    rsTabela.Close

    'Calcula Saldo Final (Total geral de Duplicatas a Receber em Aberto)
    lSQL = ""
    lSQL = lSQL & "SELECT SUM([Valor do Vencimento]) As Total"
    lSQL = lSQL & "  FROM Duplicata_Receber"
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        If Not IsNull(rsTabela("Total").Value) Then
            lSaldoFinal = lSaldoFinal + rsTabela("Total").Value
        End If
    End If
    rsTabela.Close

    'Calcula Saldo Inical
    lSaldoInicial = lSaldoFinal - lTotalEntrada + lTotalSaida

End Sub
Private Sub Relatorio()
    ZeraVariaveis
    CalculoInicial
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Codigo, [Razao Social] AS NomeCliente "
    lSQL = lSQL & "  FROM Cliente"
    lSQL = lSQL & " ORDER BY [Razao Social]"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        ImpDados
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
    If lPagina > 0 Then
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Análise das Notas de Abastecimento|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub ImpDados()
    Dim xLinha As String
    Dim xEntrada As Currency
    Dim xSaida As Currency
    Dim xSaldo As Currency
    
    xSaldo = lSaldoInicial
    Do Until rsTabela.EOF
        xEntrada = 0
        xSaida = 0
        If lPagina = 0 Then
            ImpCab
        End If
        If lLinha >= 55 Then
            xLinha = "+---------------------------------+--------------+--------------+--------------+"
            Mid(xLinha, 25, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & xLinha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        
        'Calcula Entradas (Notas em Aberto no Período)
        lSQL = ""
        lSQL = lSQL & "SELECT SUM([Valor Total]) As Total"
        lSQL = lSQL & "  FROM Movimento_Nota_Abastecimento"
        lSQL = lSQL & " WHERE [Codigo do Cliente] = " & rsTabela("Codigo").Value
        lSQL = lSQL & "   AND [Data do Abastecimento] >= " & preparaData(msk_data_i.Text)
        lSQL = lSQL & "   AND [Data do Abastecimento] <= " & preparaData(msk_data_f.Text)
        Set rsTabela2 = New adodb.Recordset
        Set rsTabela2 = Conectar.RsConexao(lSQL)
        'Verifica movimento
        If rsTabela2.RecordCount > 0 Then
            If Not IsNull(rsTabela2("Total").Value) Then
                xEntrada = xEntrada + rsTabela2("Total").Value
            End If
        End If
        rsTabela2.Close
    
        'Calcula Entradas (Notas Baixadas no Período)
        lSQL = ""
        lSQL = lSQL & "SELECT SUM([Valor Total]) As Total"
        lSQL = lSQL & "  FROM Baixa_Nota_Abastecimento"
        lSQL = lSQL & " WHERE [Codigo do Cliente] = " & rsTabela("Codigo").Value
        lSQL = lSQL & "   AND [Data do Abastecimento] >= " & preparaData(msk_data_i.Text)
        lSQL = lSQL & "   AND [Data do Abastecimento] <= " & preparaData(msk_data_f.Text)
        Set rsTabela2 = New adodb.Recordset
        Set rsTabela2 = Conectar.RsConexao(lSQL)
        'Verifica movimento
        If rsTabela2.RecordCount > 0 Then
            If Not IsNull(rsTabela2("Total").Value) Then
                xEntrada = xEntrada + rsTabela2("Total").Value
            End If
        End If
        rsTabela2.Close
        
        'Calcula Saídas (Duplicatas Recebidas no Período
        lSQL = ""
        lSQL = lSQL & "SELECT SUM([Valor Pago]) As Total"
        lSQL = lSQL & "  FROM Baixa_Duplicata_Receber"
        lSQL = lSQL & " WHERE [Codigo do Cliente] = " & rsTabela("Codigo").Value
        lSQL = lSQL & "   AND [Data do Pagamento] >= " & preparaData(msk_data_i.Text)
        lSQL = lSQL & "   AND [Data do Pagamento] <= " & preparaData(msk_data_f.Text)
        Set rsTabela2 = New adodb.Recordset
        Set rsTabela2 = Conectar.RsConexao(lSQL)
        'Verifica movimento
        If rsTabela2.RecordCount > 0 Then
            If Not IsNull(rsTabela2("Total").Value) Then
                xSaida = xSaida + rsTabela2("Total").Value
            End If
        End If
        rsTabela2.Close
        
        If xEntrada > 0 Then
            xSaldo = xSaldo + xEntrada - xSaida
            Call ImpDet(xEntrada, xSaida, xSaldo)
        End If
        rsTabela.MoveNext
    Loop
    'If lTotal > 0 Then
        ImpTotal
    'End If
End Sub
Private Sub ImpDet(ByVal pEntrada As Currency, ByVal pSaida As Currency, ByVal pSaldo As Currency)
    Dim xLinha As String
    Dim i As Integer
    
    '                                         1         2         3         4         5         6         7         8
    '                                12345678901234567890123456789012345678901234567890123456789012345678901234567890
    '                                cliente                           |    95.471,22 |    82.701,65 | 1.102.957,59 |
    'BioImprime "@Printer.Print " & "+---------------------------------+--------------+--------------+--------------+"
    'BioImprime "@Printer.Print " & "|RAZAO SOCIAL                     |    ENTRADA   |     SAIDA    |     SALDO    |"
    'BioImprime "@Printer.Print " & "+---------------------------------+--------------+--------------+--------------+"
    xLinha = "|                                 |              |              |              |"
    Mid(xLinha, 2, 33) = rsTabela("NomeCliente").Value
    i = Len(Format(pEntrada, "##,###,##0.00"))
    Mid(xLinha, 36 + 13 - i, i) = Format(pEntrada, "##,###,##0.00")
    i = Len(Format(pSaida, "##,###,##0.00"))
    Mid(xLinha, 51 + 13 - i, i) = Format(pSaida, "##,###,##0.00")
    i = Len(Format(pSaldo, "##,###,##0.00"))
    Mid(xLinha, 66 + 13 - i, i) = Format(pSaldo, "##,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    xLinha = "+------------------+--------------+--------------+--------------+--------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                  |              |              |              |              |"
    i = Len(Format(lSaldoInicial, "##,###,##0.00"))
    Mid(xLinha, 21 + 13 - i, i) = Format(lSaldoInicial, "##,###,##0.00")
    i = Len(Format(lTotalEntrada, "##,###,##0.00"))
    Mid(xLinha, 36 + 13 - i, i) = Format(lTotalEntrada, "##,###,##0.00")
    i = Len(Format(lTotalSaida, "##,###,##0.00"))
    Mid(xLinha, 51 + 13 - i, i) = Format(lTotalSaida, "##,###,##0.00")
    i = Len(Format(lSaldoFinal, "##,###,##0.00"))
    Mid(xLinha, 66 + 13 - i, i) = Format(lSaldoFinal, "##,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+------------------+--------------+--------------+--------------+--------------+"
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
    BioImprime "@Printer.Print " & "+------------------------------------------------------------------------------+"
    x_string_40 = g_nome_empresa
    BioImprime "@Printer.Print " & "| " & x_string_40 & "                         Página, " & Format(lPagina, "000") & " |"
    xLinha = "| ANALISE DAS NOTAS DE ABASTECIMENTO                        CIDADE, __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@Printer.Print " & "| Referente a.: " & msk_data_i.Text & " a " & msk_data_f.Text & "                                        |"
    BioImprime "@@Printer.FontBold = False"
    '                                        1         2         3         4         5         6         7         8
    '                               12345678901234567890123456789012345678901234567890123456789012345678901234567890
    '                               cliente                           |    95.471,22 |    82.701,65 | 1.102.957,59 |
    BioImprime "@Printer.Print " & "+---------------------------------+--------------+--------------+--------------+"
    BioImprime "@Printer.Print " & "|RAZAO SOCIAL                     |    ENTRADA   |     SAIDA    |     SALDO    |"
    BioImprime "@Printer.Print " & "+---------------------------------+--------------+--------------+--------------+"
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
