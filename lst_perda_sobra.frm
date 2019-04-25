VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_perda_sobra 
   Caption         =   "Perda x Sobra do L.M.C."
   ClientHeight    =   3075
   ClientLeft      =   2790
   ClientTop       =   3810
   ClientWidth     =   5475
   Icon            =   "lst_perda_sobra.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_perda_sobra.frx":030A
   ScaleHeight     =   3075
   ScaleWidth      =   5475
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   840
      Picture         =   "lst_perda_sobra.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Visualiza perda/sobra do L.M.C."
      Top             =   2160
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2340
      Picture         =   "lst_perda_sobra.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Imprime perda/sobra do L.M.C."
      Top             =   2160
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3840
      Picture         =   "lst_perda_sobra.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2160
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5235
      Begin VB.CheckBox chkAcumulaDiferenca 
         Caption         =   "A&cumula Diferença"
         Height          =   255
         Left            =   2700
         TabIndex        =   12
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin VB.TextBox txtBico 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   11
         Top             =   1500
         Width           =   435
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_perda_sobra.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_perda_sobra.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_perda_sobra.frx":6CBA
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
         Width           =   975
         _ExtentX        =   1720
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
         Width           =   975
         _ExtentX        =   1720
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
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Código do &Bico"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1500
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
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1080
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
      Top             =   2580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_perda_sobra"
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
Dim lData As Date
Dim lQuantidadeBomba As Integer
Dim lEncerranteLMC(0 To gQUANTIDADE_MAXIMA_BICO) As Currency
Dim lObservacao(0 To gQUANTIDADE_MAXIMA_BICO) As String
Dim iObs As Integer
Dim lDiferencaDeduzir As Currency
Dim lSQL As String

Dim rstCombustivel As New ADODB.Recordset
Dim rstMovimentoBomba As New ADODB.Recordset
Dim rstMovimentoBombaLMC As New ADODB.Recordset
Dim rstMovimentoBombaLMC_QTD As New ADODB.Recordset

Private Bomba As New cBomba
Private Configuracao As New cConfiguracao
Private Sub AdcionaMensagem(ByVal pMensagem As String)
    If iObs < 30 Then
        lObservacao(iObs) = pMensagem
        iObs = iObs + 1
    ElseIf iObs = 30 Then
        If lObservacao(30) = "" Then
            lObservacao(30) = pMensagem
        Else
            lObservacao(30) = "Estourou "
        End If
    End If
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
Private Sub BuscaDadosCombustivel()
    lSQL = ""
    lSQL = lSQL & "SELECT Codigo, Nome"
    lSQL = lSQL & " FROM Combustivel"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " ORDER BY Codigo, Nome"
    Set rstCombustivel = New ADODB.Recordset
    Set rstCombustivel = Conectar.RsConexao(lSQL)
End Sub
Private Sub BuscaDadosMovimentoBomba()
    lSQL = ""
    lSQL = lSQL & "SELECT Movimento_Bomba.[Codigo da Bomba], Movimento_Bomba.Data, Movimento_Bomba.Encerrante"
    lSQL = lSQL & "  FROM movimento_bomba"
    lSQL = lSQL & "    RIGHT OUTER JOIN"
    lSQL = lSQL & "    ("
    lSQL = lSQL & "       SELECT Movimento_Bomba.Empresa, Movimento_Bomba.[Codigo da Bomba], Movimento_Bomba.Data, MAX(Movimento_Bomba.Periodo) AS Periodo, MAX(Movimento_Bomba.SubCaixa) AS SubCaixa"
    lSQL = lSQL & "         FROM movimento_bomba"
    lSQL = lSQL & "        WHERE Empresa = " & g_empresa
    lSQL = lSQL & "          AND Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "          AND Data <= " & preparaData(msk_data_f.Text)
    If txtBico.Text <> "" Then
        lSQL = lSQL & "      AND [Codigo da Bomba] = " & txtBico.Text
    End If
    lSQL = lSQL & "        GROUP BY Empresa, [Codigo da Bomba], Data"
    lSQL = lSQL & "    ) AS SQ On SQ.Empresa = Movimento_Bomba.Empresa AND SQ.[Codigo da Bomba] = Movimento_Bomba.[Codigo da Bomba] AND SQ.Data = Movimento_Bomba.Data AND SQ.Periodo = Movimento_Bomba.Periodo AND SQ.SubCaixa = Movimento_Bomba.SubCaixa"
    lSQL = lSQL & "    ORDER BY Movimento_Bomba.[Codigo da Bomba], Movimento_Bomba.Data"
    Set rstMovimentoBomba = New ADODB.Recordset
    Set rstMovimentoBomba = Conectar.RsConexao(lSQL)
End Sub
Private Sub BuscaDadosMovimentoBombaLMC()
    lSQL = ""
    lSQL = lSQL & "SELECT Movimento_Bomba_LMC.[Codigo da Bomba], Movimento_Bomba_LMC.Data, Movimento_Bomba_LMC.Encerrante, Movimento_Bomba_LMC.[Tipo de Combustivel]"
    lSQL = lSQL & "  FROM Movimento_Bomba_LMC"
    lSQL = lSQL & "    RIGHT OUTER JOIN"
    lSQL = lSQL & "    ("
    lSQL = lSQL & "       SELECT Movimento_Bomba_LMC.Empresa, Movimento_Bomba_LMC.[Codigo da Bomba], Movimento_Bomba_LMC.Data, MAX(Movimento_Bomba_LMC.Periodo) AS Periodo, MAX(Movimento_Bomba_LMC.SubCaixa) AS SubCaixa"
    lSQL = lSQL & "         FROM Movimento_Bomba_LMC"
    lSQL = lSQL & "        WHERE Empresa = " & g_empresa
    lSQL = lSQL & "          AND Data >= " & preparaData(CDate(msk_data_i.Text) - 1)
    lSQL = lSQL & "          AND Data <= " & preparaData(msk_data_f.Text)
    If txtBico.Text <> "" Then
        lSQL = lSQL & "      AND [Codigo da Bomba] = " & txtBico.Text
    End If
    lSQL = lSQL & "        GROUP BY Empresa, [Codigo da Bomba], Data"
    lSQL = lSQL & "    ) AS SQ On SQ.Empresa = Movimento_Bomba_LMC.Empresa AND SQ.[Codigo da Bomba] = Movimento_Bomba_LMC.[Codigo da Bomba] AND SQ.Data = Movimento_Bomba_LMC.Data AND SQ.Periodo = Movimento_Bomba_LMC.Periodo AND SQ.SubCaixa = Movimento_Bomba_LMC.SubCaixa"
    lSQL = lSQL & "    ORDER BY Movimento_Bomba_LMC.[Codigo da Bomba], Movimento_Bomba_LMC.Data"
    Set rstMovimentoBombaLMC = New ADODB.Recordset
    Set rstMovimentoBombaLMC = Conectar.RsConexao(lSQL)
End Sub
Private Sub BuscaDadosMovimentoBombaLMC_QTD()
    lSQL = ""
    lSQL = lSQL & "SELECT Movimento_Bomba_LMC.[Codigo da Bomba], Movimento_Bomba_LMC.Data, SUM(Movimento_Bomba_LMC.[Quantidade da Saida]) AS TotalQuantidade"
    lSQL = lSQL & "  FROM Movimento_Bomba_LMC"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data >= " & preparaData(CDate(msk_data_i.Text) - 1)
    lSQL = lSQL & "   AND Data <= " & preparaData(msk_data_f.Text)
    If txtBico.Text <> "" Then
        lSQL = lSQL & "      AND [Codigo da Bomba] = " & txtBico.Text
    End If
    lSQL = lSQL & " GROUP BY [Codigo da Bomba], Data"
    lSQL = lSQL & " ORDER BY [Codigo da Bomba], Data"
    Set rstMovimentoBombaLMC_QTD = New ADODB.Recordset
    Set rstMovimentoBombaLMC_QTD = Conectar.RsConexao(lSQL)
End Sub
Private Function BuscaTipoCombustivel(ByVal pTipoCombustivel As String) As String
    BuscaTipoCombustivel = "** Inexistente: " & pTipoCombustivel & " **"
    rstCombustivel.MoveFirst
    rstCombustivel.Find "Codigo = " & preparaTexto(pTipoCombustivel)
    If Not rstCombustivel.EOF Then
        BuscaTipoCombustivel = rstCombustivel!Nome
    End If
End Function
Private Function BuscaRegistroMovimetoBomba(ByVal pCodigoBomba As String, ByVal pData As Date) As Boolean
    Dim xCondicao As String
    xCondicao = " [Codigo da Bomba] = " & preparaTexto(pCodigoBomba)
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    
    BuscaRegistroMovimetoBomba = False
    rstMovimentoBomba.Filter = ""
    If rstMovimentoBomba.RecordCount > 0 Then
        rstMovimentoBomba.MoveFirst
        rstMovimentoBomba.Filter = xCondicao
        If Not rstMovimentoBomba.EOF Then
            BuscaRegistroMovimetoBomba = True
        End If
    End If
End Function
Private Function BuscaRegistroMovimetoBombaLMC(ByVal pCodigoBomba As String, ByVal pData As Date) As Boolean
    Dim xCondicao As String
    xCondicao = " [Codigo da Bomba] = " & preparaTexto(pCodigoBomba)
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    
    BuscaRegistroMovimetoBombaLMC = False
    rstMovimentoBombaLMC.Filter = ""
    If rstMovimentoBombaLMC.RecordCount > 0 Then
        rstMovimentoBombaLMC.MoveFirst
        rstMovimentoBombaLMC.Filter = xCondicao
        If Not rstMovimentoBombaLMC.EOF Then
            BuscaRegistroMovimetoBombaLMC = True
        End If
    End If
End Function
Private Function BuscaRegistroMovimetoBombaLMC_QTD(ByVal pCodigoBomba As String, ByVal pData As Date) As Boolean
    Dim xCondicao As String
    xCondicao = " [Codigo da Bomba] = " & pCodigoBomba
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    
    BuscaRegistroMovimetoBombaLMC_QTD = False
    rstMovimentoBombaLMC_QTD.Filter = ""
    If rstMovimentoBombaLMC_QTD.RecordCount > 0 Then
        rstMovimentoBombaLMC_QTD.MoveFirst
        rstMovimentoBombaLMC_QTD.Filter = xCondicao
        If Not rstMovimentoBombaLMC_QTD.EOF Then
            BuscaRegistroMovimetoBombaLMC_QTD = True
        End If
    End If
End Function
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Bomba = Nothing
    Set Configuracao = Nothing
    
    Set rstCombustivel = Nothing
    Set rstMovimentoBomba = Nothing
    Set rstMovimentoBombaLMC = Nothing
    Set rstMovimentoBombaLMC_QTD = Nothing
End Sub
Private Sub ImprimeMensagem()
    Dim xLinha As String
    Dim i As Integer

    For i = 0 To iObs - 1
        xLinha = lObservacao(i)
        BioImprime "@Printer.Print " & xLinha
    Next
End Sub
Private Sub ImpCab()
    Dim x_linha As String
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
    x_linha = "+------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|                                                                  Página: ___ |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    Mid(x_linha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| PERDAS X SOBRAS DO L.M.C.                                Goiânia, __/__/____ |"
    Mid(x_linha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| REFERENTE A.: __/__/____ A __/__/____                                        |"
    Mid(x_linha, 17, 10) = msk_data_i.Text
    Mid(x_linha, 30, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    '                   1         2         3         4         5         6         7        80
    '          12345678901234567890123456789012345678901234567890123456789012345678901234567890
    x_linha = "+----------+----+-----------------+--------------+--------------+--------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|   DATA   |COD.| COMBUSTIVEL     |  ENCERRANTE  |  ENCERRANTE  |   DIFERENCA  |"
'              |01/01/2004|  1 | Gasolina Comum  |   405.905,30 |   606.699,60 |  -200.794,30 |
    
    
    
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| MOVIMENTO|BICO|                 |  DO  L.M.C.  |  DA   BOMBA  |              |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "+----------+----+-----------------+--------------+--------------+--------------+"
    BioImprime "@Printer.Print " & x_linha
End Sub
Private Sub ImpDet(ByVal pData As Date, ByVal pBomba As Integer, ByVal pCombustivel As String, ByVal pEncerranteLMC As Currency, ByVal pEncerrante As Currency, ByVal pDiferenca As Currency)
    Dim xLinha As String
    Dim i As Integer

    If lPagina = 0 Then
        ImpCab
    End If
    If lLinha >= 60 Then
        xLinha = "+----------+----+-----------------+--------------+--------------+--------------+"
        Mid(xLinha, 5, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    '                  1         2         3         4         5         6         7        80
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "|          |    |                 |              |              |              |"
    Mid(xLinha, 2, 10) = Format(pData, "dd/mm/yyyy")
    i = Len(Format(pBomba, "#0"))
    Mid(xLinha, 14 + 2 - i, i) = Format(pBomba, "#0")
    Mid(xLinha, 19, 15) = pCombustivel
    i = Len(Format(pEncerranteLMC, "#####,##0.00"))
    Mid(xLinha, 37 + 12 - i, i) = Format(pEncerranteLMC, "#####,##0.00")
    i = Len(Format(pEncerrante, "#####,##0.00"))
    Mid(xLinha, 52 + 12 - i, i) = Format(pEncerrante, "#####,##0.00")
    If pDiferenca <> 0 Then
        i = Len(Format(pDiferenca, "#####,##0.00"))
        Mid(xLinha, 67 + 12 - i, i) = Format(pDiferenca, "#####,##0.00")
    End If
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpDetErro(ByVal pData As Date, ByVal pDiferenca As Currency)
    Dim xLinha As String
    Dim i As Integer

    '                  1         2         3         4         5         6         7        80
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "** ERRO DE CONTINUIDADE DE ENCERRANTE                               LTS ********"
    If pDiferenca < 0 Then
        Mid(xLinha, 39, 17) = " PASSOU VENDA DE "
        AdcionaMensagem (Format(pData, "dd/mm/yyyy") & ": ENCERRANTE VOLTOU " & Format(pDiferenca, "#####,##0.00")) & " LITROS."
    Else
        Mid(xLinha, 39, 17) = " FALTOU VENDA DE "
        AdcionaMensagem (Format(pData, "dd/mm/yyyy") & ": ENCERRANTE ADIANTOU " & Format(pDiferenca, "#####,##0.00")) & " LITROS."
    End If
    i = Len(Format(pDiferenca, "#####,##0.00"))
    Mid(xLinha, 56 + 12 - i, i) = Format(pDiferenca, "#####,##0.00")
    
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    xLinha = "+----------+----+-----------------+--------------+--------------+--------------+"
    Mid(xLinha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    ImprimeMensagem
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub Relatorio()
    BuscaDadosCombustivel
    BuscaDadosMovimentoBomba
    BuscaDadosMovimentoBombaLMC
    BuscaDadosMovimentoBombaLMC_QTD
    ZeraVariaveis
    lData = CDate(msk_data_i.Text)
    'Loop data
    Do Until lData > CDate(msk_data_f.Text)
        Call LoopData(lData)
        lData = lData + 1
    Loop
    ImpTotal
    BioImprime "@@Printer.EndDoc"
    BioFechaImprime
    g_string = lLocal & lNomeArquivo & "|@|Relatório das Perdas x Sobras do LMC|@|"
    frm_preview.Show 1
End Sub
Private Sub chkAcumulaDiferenca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        txtBico.SetFocus
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
    txtBico.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        txtBico.SetFocus
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
            Call GravaAuditoria(1, Me.name, 7, "Ref:" & msk_data_i.Text & " a " & msk_data_f.Text & " Bico:" & txtBico.Text)
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
    ElseIf txtBico.Text <> "" Then
        If Bomba.LocalizarCodigo(g_empresa, Val(txtBico.Text)) Then
            ValidaCampos = True
        Else
            MsgBox "Nao existe Bico/Bomba cadastrado com este código.", vbInformation, "Atenção!"
            txtBico.SetFocus
        End If
    Else
        ValidaCampos = True
    End If
End Function
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    If Configuracao.LocalizarCodigo(g_empresa) Then
        lQuantidadeBomba = Configuracao.QuantidadeBico
    End If
    For iObs = 0 To gQUANTIDADE_MAXIMA_BICO
        lEncerranteLMC(iObs) = 0
    Next
    For iObs = 0 To gQUANTIDADE_MAXIMA_BICO
        lObservacao(iObs) = ""
    Next
    lDiferencaDeduzir = 0
    iObs = 0
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    lLocal = 0
    If ValidaCampos Then
        AtivaBotoes (False)
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 6, "Ref:" & msk_data_i.Text & " a " & msk_data_f.Text & " Bico:" & txtBico.Text)
            Relatorio
        End If
        AtivaBotoes (True)
        txtBico.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        msk_data_i.Text = CDate(fDataPrimeiroDiaMesAnterior(Date)) - 1
        msk_data_f.Text = fDataUltimoDiaMesAnterior(Date)
        msk_data_i.SetFocus
    End If
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
    Screen.MousePointer = 1
    CentraForm Me
End Sub
Private Sub LoopData(ByVal pData As Date)
    Dim xLinha As String
    Dim xEncerrante As Currency
    Dim xEncerranteLMC As Currency
    Dim xDiferenca As Currency
    Dim xDiferenca2 As Currency
    Dim xCombustivel As String
    Dim xQtdVenda As Currency
    Dim i As Integer
    Dim i2 As Integer
    Dim xExisteMovBombaLMC As Boolean
    
    xDiferenca2 = 0
    i2 = 1
    If txtBico.Text <> "" Then
        i2 = Val(txtBico.Text)
        lQuantidadeBomba = Val(txtBico.Text)
    End If
    For i = i2 To lQuantidadeBomba
        If lEncerranteLMC(i) = 0 Then
            'If MovimentoBombaLMC.LocalizarUltimoPeriodoBico(g_empresa, CDate(pData - 1), i, 0) Then
            '    lEncerranteLMC(i) = MovimentoBombaLMC.Encerrante
            'End If
            If BuscaRegistroMovimetoBombaLMC(i, pData - 1) Then
                lEncerranteLMC(i) = rstMovimentoBombaLMC!Encerrante
            End If
        End If
        xEncerrante = 0
        xEncerranteLMC = 0
        'If MovimentoBomba.LocalizarUltimoPeriodoBico(g_empresa, pData, i, 0) Then
        '    xEncerrante = MovimentoBomba.Encerrante
        'End If
        If BuscaRegistroMovimetoBomba(i, pData) Then
            xEncerrante = rstMovimentoBomba!Encerrante
        End If
        'If MovimentoBombaLMC.LocalizarUltimoPeriodoBico(g_empresa, pData, i, 0) Then
        '    xEncerranteLMC = MovimentoBombaLMC.Encerrante
        'End If
        xExisteMovBombaLMC = False
        If BuscaRegistroMovimetoBombaLMC(i, pData) Then
            xEncerranteLMC = rstMovimentoBombaLMC!Encerrante
            xExisteMovBombaLMC = True
        End If
        xDiferenca = xEncerranteLMC - xEncerrante
        If chkAcumulaDiferenca.Value = 0 Then
            If xDiferenca <> 0 Then
                If lDiferencaDeduzir <> 0 Then
                    xDiferenca2 = xDiferenca
                    If xDiferenca = lDiferencaDeduzir Then
                        xDiferenca = 0
                    Else
                        xDiferenca = xDiferenca2 - lDiferencaDeduzir
                        lDiferencaDeduzir = lDiferencaDeduzir + xDiferenca
                    End If
                Else
                    lDiferencaDeduzir = xDiferenca
                End If
            End If
        End If
        'xCombustivel = ""
        'If Combustivel.LocalizarCodigo(g_empresa, MovimentoBombaLMC.TipoCombustivel) Then
        '    xCombustivel = Combustivel.Nome
        'End If
        If xExisteMovBombaLMC Then
            xCombustivel = BuscaTipoCombustivel(rstMovimentoBombaLMC![Tipo de Combustivel])
        Else
            xCombustivel = "SEM MOV. LMC"
        End If
        'If xEncerranteLMC > 0 And xEncerrante > 0 Then
            Call ImpDet(pData, i, xCombustivel, xEncerranteLMC, xEncerrante, xDiferenca)
        'End If
        'xQtdVenda = MovimentoBombaLMC.TotalLitrosBicoPeriodo(g_empresa, pData, pData, i, 1, 9, "")
        xQtdVenda = 0
        If BuscaRegistroMovimetoBombaLMC_QTD(i, pData) Then
            xQtdVenda = rstMovimentoBombaLMC_QTD!TotalQuantidade
        End If
        If xEncerranteLMC <> (lEncerranteLMC(i) + xQtdVenda) Then
            xQtdVenda = xEncerranteLMC - (lEncerranteLMC(i) + xQtdVenda)
            Call ImpDetErro(pData, xQtdVenda)
        End If
        If xEncerranteLMC = 0 Then
            AdcionaMensagem (Format(pData, "dd/mm/yyyy") & ": FALTA MOVIMENTO DE BOMBA DO LMC NESTA DATA")
        End If
        lEncerranteLMC(i) = xEncerranteLMC
    Next
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
        txtBico.SetFocus
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
Private Sub txtBico_GotFocus()
    txtBico.SelStart = 0
    txtBico.SelLength = Len(txtBico.Text)
End Sub
Private Sub txtBico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
