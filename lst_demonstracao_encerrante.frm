VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form lst_demonstracao_encerrante 
   Caption         =   "Demonstração dos Encerrantes"
   ClientHeight    =   2760
   ClientLeft      =   2790
   ClientTop       =   3810
   ClientWidth     =   5475
   Icon            =   "lst_demonstracao_encerrante.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_demonstracao_encerrante.frx":030A
   ScaleHeight     =   2760
   ScaleWidth      =   5475
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   840
      Picture         =   "lst_demonstracao_encerrante.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Visualiza Demonstrativo dos Encerrantes."
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2340
      Picture         =   "lst_demonstracao_encerrante.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprime Demonstrativo dos Encerrantes."
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3840
      Picture         =   "lst_demonstracao_encerrante.frx":3074
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
      Width           =   5235
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_demonstracao_encerrante.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_demonstracao_encerrante.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_demonstracao_encerrante.frx":6CBA
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
      Top             =   2220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_demonstracao_encerrante"
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

Private Bomba As New cBomba
Private Combustivel As New cCombustivel
Private Configuracao As New cConfiguracao
Private MedicaoCombustivel As New cMedicaoCombustivel
Private MovimentoBomba As New cMovimentoBomba
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
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Bomba = Nothing
    Set Combustivel = Nothing
    Set Configuracao = Nothing
    Set MedicaoCombustivel = Nothing
    Set MovimentoBomba = Nothing
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
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.CurrentY = 0"
    x_linha = "+------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|                                                                  Página: ___ |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    Mid(x_linha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| DEMONSTRATIVO DOS ENCERRANTES                            Goiânia, __/__/____ |"
    If g_nome_usuario = "L.M.C." Then
        Mid(x_linha, 33, 10) = "DO LMC    "
    Else
        Mid(x_linha, 33, 10) = "DAS BOMBAS"
    End If
    Mid(x_linha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| REFERENTE A.: __/__/____ A __/__/____                                        |"
    Mid(x_linha, 17, 10) = msk_data_i.Text
    Mid(x_linha, 30, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| TODAS AS BOMBAS SÃO DA MARCA:                                                |"
    If UCase(g_nome_empresa) Like "*MANTIQUEIRA*" Then
        Mid(x_linha, 33, 10) = "WAYNE     "
    End If
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    '                   1         2         3         4         5         6         7        80
    '          12345678901234567890123456789012345678901234567890123456789012345678901234567890
    x_linha = "+----------+----+-----------------+--------------+--------------+--------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|  NUMERO  |COD.| COMBUSTIVEL     |   ABERTURA   |  ENCERRANTE  |    LITROS    |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| DE SERIE |BICO|                 |  __/__/____  |  __/__/____  |   VENDIDOS   |"
    Mid(x_linha, 38, 10) = msk_data_i.Text
    Mid(x_linha, 53, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & x_linha
    x_linha = "+----------+----+-----------------+--------------+--------------+--------------+"
    BioImprime "@Printer.Print " & x_linha
End Sub
Private Sub ImpDet(ByVal pBomba As Integer, ByVal pCombustivel As String, ByVal pAbertura As Currency, ByVal pEncerrante As Currency, ByVal pLitrosVendidos As Currency)
    Dim xLinha As String
    Dim xQuantidade As Currency
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
    If Bomba.LocalizarCodigo(g_empresa, pBomba) Then
        Mid(xLinha, 3, 9) = Bomba.NumeroSerie
    End If
    i = Len(Format(pBomba, "#0"))
    Mid(xLinha, 14 + 2 - i, i) = Format(pBomba, "#0")
    Mid(xLinha, 19, 15) = pCombustivel
    i = Len(Format(pAbertura, "#####,##0.00"))
    Mid(xLinha, 37 + 12 - i, i) = Format(pAbertura, "#####,##0.00")
    i = Len(Format(pEncerrante, "#####,##0.00"))
    Mid(xLinha, 52 + 12 - i, i) = Format(pEncerrante, "#####,##0.00")
    i = Len(Format(pLitrosVendidos, "#####,##0.00"))
    Mid(xLinha, 67 + 12 - i, i) = Format(pLitrosVendidos, "#####,##0.00")
    xQuantidade = pEncerrante - pAbertura
    If xQuantidade <> pLitrosVendidos Then
        Mid(xLinha, 79, 2) = "* "
    End If
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    xLinha = "+----------+----+-----------------+--------------+--------------+--------------+"
    Mid(xLinha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
End Sub
Private Sub LoopCombustivel()
    Dim xLinha As String
    Dim i As Integer
    xLinha = " "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@Printer.Print " & xLinha
    xLinha = "         +----------------------+--------+--------------+         "
    BioImprime "@Printer.Print " & xLinha
    xLinha = "         | COMBUSTÍVEL          | TANQUE |  QUANTIDADE  |         "
    BioImprime "@Printer.Print " & xLinha
    xLinha = "         +----------------------+--------+--------------+         "
    BioImprime "@Printer.Print " & xLinha
    If Combustivel.LocalizarPrimeiro(g_empresa) Then
        LoopMedidaTanque
        Do Until Combustivel.LocalizarProximo = False
            LoopMedidaTanque
        Loop
    End If
    xLinha = "         +----------------------+--------+--------------+         "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub LoopMedidaTanque()
    'Localiza Medição de Combustível de Fechamento do Dia
    If MedicaoCombustivel.LocalizarPrimeiroTanqueComb(g_empresa, CDate(msk_data_f) + 1, Combustivel.Codigo) Then
        ImpMedidaTanque
        Do Until MedicaoCombustivel.LocalizarProximoTanqueComb(g_empresa, CDate(msk_data_f) + 1, Combustivel.Codigo) = False
            ImpMedidaTanque
        Loop
    End If
End Sub
Private Sub ImpMedidaTanque()
    Dim xLinha As String
    Dim i As Integer
    '                  1         2         3         4         5         6"
    '         123456789012345678901234567890123456789012345678901234567890"
    xLinha = "         |                      |        |              |         "
    
    
    Mid(xLinha, 12, 20) = Combustivel.Nome
    Mid(xLinha, 37, 2) = Format(MedicaoCombustivel.NumeroTanque, "00")
    i = Len(Format(MedicaoCombustivel.Quantidade, "#####,##0.00"))
    Mid(xLinha, 44 + 12 - i, i) = Format(MedicaoCombustivel.Quantidade, "#####,##0.00")
    BioImprime "@Printer.Print " & xLinha
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    lData = CDate(msk_data_i.Text)
    Call LeEncerrantes
    ImpTotal
    LoopCombustivel
    BioImprime "@@Printer.EndDoc"
    BioFechaImprime
    g_string = lLocal & lNomeArquivo & "|@|Relatório da Demonstração dos Encerrantes|@|"
    frm_preview.Show 1
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cmd_visualizar.SetFocus
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
    cmd_visualizar.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
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
        AtivaBotoes (False)
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "Ref:" & msk_data_i.Text & " a " & msk_data_f.Text)
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
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    If Configuracao.LocalizarCodigo(g_empresa) Then
        lQuantidadeBomba = Configuracao.QuantidadeBico
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    lLocal = 0
    If ValidaCampos Then
        AtivaBotoes (False)
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 6, "Ref:" & msk_data_i.Text & " a " & msk_data_f.Text)
            Relatorio
        End If
        AtivaBotoes (True)
        cmd_sair.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Dim xString As String
    
    If g_nome_usuario = "L.M.C." Then
        xString = "LMC"
    Else
        xString = ""
    End If
    Call GravaAuditoria(1, Me.name, 1, xString)
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        msk_data_i.Text = fDataPrimeiroDiaMesAnterior(Date)
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
    
    If g_nome_usuario = "L.M.C." Then
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivelLMC"
        MovimentoBomba.NomeTabela = "Movimento_Bomba_LMC"
    Else
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivel"
        MovimentoBomba.NomeTabela = "Movimento_Bomba"
    End If
End Sub
Private Sub LeEncerrantes()
    Dim xLinha As String
    Dim xAbertura As Currency
    Dim xEncerrante As Currency
    Dim xLitrosVendidos As Currency
    Dim xCombustivel As String
    Dim i As Integer
    For i = 1 To lQuantidadeBomba
        xAbertura = 0
        xEncerrante = 0
        If MovimentoBomba.LocalizarPrimeiroPeriodoBico(g_empresa, CDate(msk_data_i.Text), i, 0) Then
            xAbertura = MovimentoBomba.Abertura
        End If
        If MovimentoBomba.LocalizarUltimoPeriodoBico(g_empresa, CDate(msk_data_f.Text), i, 0) Then
            xEncerrante = MovimentoBomba.Encerrante
        End If
        xLitrosVendidos = MovimentoBomba.TotalLitrosBicoPeriodo(g_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), i, 1, 9, "")
        xCombustivel = ""
        If Combustivel.LocalizarCodigo(g_empresa, MovimentoBomba.TipoCombustivel) Then
            xCombustivel = Combustivel.Nome
        End If
        Call ImpDet(i, xCombustivel, xAbertura, xEncerrante, xLitrosVendidos)
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
