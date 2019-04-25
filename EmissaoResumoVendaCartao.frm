VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form EmissaoResumoVendaCartao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumo das Vendas de Cartão de Crédito"
   ClientHeight    =   3750
   ClientLeft      =   2775
   ClientTop       =   3795
   ClientWidth     =   5475
   Icon            =   "EmissaoResumoVendaCartao.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "EmissaoResumoVendaCartao.frx":030A
   ScaleHeight     =   3750
   ScaleWidth      =   5475
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   840
      Picture         =   "EmissaoResumoVendaCartao.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Visualiza resumo do L.M.C."
      Top             =   2820
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2340
      Picture         =   "EmissaoResumoVendaCartao.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprime resumo do L.M.C."
      Top             =   2820
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3840
      Picture         =   "EmissaoResumoVendaCartao.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2820
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5235
      Begin VB.OptionButton optVencimento 
         Caption         =   "Vencimento"
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton optEmissao 
         Caption         =   "Emissão"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   1920
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CheckBox chkLinhaSeparadora 
         Caption         =   "Imprime linha separadora"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2220
         Width           =   2235
      End
      Begin VB.ComboBox cboCartaoCredito 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1500
         Width           =   3435
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "EmissaoResumoVendaCartao.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2700
         Picture         =   "EmissaoResumoVendaCartao.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2700
         Picture         =   "EmissaoResumoVendaCartao.frx":6CBA
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
      Begin VB.Label Label7 
         Caption         =   "I&mprimir por"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Cartão de Crédito"
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
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "EmissaoResumoVendaCartao"
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
Dim lCodigoCartao As Integer
Dim lTotalBruto As Currency
Dim lTotalLiquido As Currency
Dim lTotalADM As Currency
Dim lSQL As String
Dim rstMovimentoBomba As New adodb.Recordset

Private CartaoCredito As New cCartaoCredito
Private MovimentoCartaoCredito As New cMovimentoCartaoCredito
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set CartaoCredito = Nothing
    Set MovimentoCartaoCredito = Nothing
End Sub
Private Sub ImpCab()
    Dim xLinha As String
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
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.CurrentY = 0"
    xLinha = "+------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = True"
    xLinha = "|                                                                  Página: ___ |"
    Mid(xLinha, 3, 40) = g_nome_empresa
    Mid(xLinha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    xLinha = "| RESUMO DAS VENDAS COM CARTÃO DE CRÉDITO                  Goiânia, __/__/____ |"
    If g_nome_usuario = "L.M.C." Then
        Mid(xLinha, 3, 40) = "RESUMO DO L.M.C.                       "
    Else
        Mid(xLinha, 3, 40) = "RESUMO DA MOVIMENTAÇÃO DE COMBUSTÍVEL  "
    End If
    Mid(xLinha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| REFERENTE A.: __/__/____ A __/__/____                                        |"
    Mid(xLinha, 17, 10) = msk_data_i.Text
    Mid(xLinha, 30, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| CARTAO......:                                                                |"
    Mid(xLinha, 17, 30) = cboCartaoCredito.Text
    BioImprime "@Printer.Print " & xLinha
    '                  1         2         3         4         5         6         7         8
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "| PERCENTUAL..: ______    PRAZO.: __ DIAS                                      |"
    Mid(xLinha, 17, 6) = "      "
    Mid(xLinha, 35, 2) = "  "
    If lCodigoCartao > 0 Then
        If CartaoCredito.LocalizarCodigo(lCodigoCartao) Then
            i = Len(Format(CartaoCredito.TaxaCusto, "##0.00"))
            Mid(xLinha, 17 + 6 - i, i) = Format(CartaoCredito.TaxaCusto, "##0.00")
            Mid(xLinha, 35, 2) = Format(CartaoCredito.DiasPrazo, "00")
        End If
    End If
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+----------+------------+------------+------------+----------------------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|   DATA   |    VALOR   |   VALOR    |    VALOR   |                            |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|          |    BRUTO   |ADMINISTRAD.|   LIQUODO  |                            |"
    If optEmissao.Value = True Then
        Mid(xLinha, 2, 10) = "  EMISSAO "
    Else
        Mid(xLinha, 2, 10) = "VENCIMENTO"
    End If
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+----------+------------+------------+------------+----------------------------+"
    BioImprime "@Printer.Print " & xLinha
End Sub
Private Sub ImpDet(ByVal pData As Date)
    Dim xLinha As String
    Dim i As Integer
    Dim xTotalBruto As Currency
    Dim xTotalLiquido As Currency
    Dim xTotalADM As Currency
    
    xTotalBruto = 0
    xTotalLiquido = 0
    xTotalADM = 0
    
    xTotalBruto = MovimentoCartaoCredito.TotalEntreDatas(g_empresa, optEmissao.Value, False, pData, pData, lCodigoCartao)
    xTotalBruto = xTotalBruto + MovimentoCartaoCredito.TotalEntreDatas(g_empresa, optEmissao.Value, True, pData, pData, lCodigoCartao)
    lTotalBruto = lTotalBruto + xTotalBruto
    
    xTotalADM = MovimentoCartaoCredito.TotalAdmEntreDatas(g_empresa, optEmissao.Value, False, pData, pData, lCodigoCartao)
    xTotalADM = xTotalADM + MovimentoCartaoCredito.TotalAdmEntreDatas(g_empresa, optEmissao.Value, True, pData, pData, lCodigoCartao)
    lTotalADM = lTotalADM + xTotalADM
    
    xTotalLiquido = xTotalBruto - xTotalADM
    lTotalLiquido = lTotalLiquido + xTotalLiquido

    If lPagina = 0 Then
        ImpCab
    End If
    If lLinha >= 60 Then
        xLinha = "+----------+------------+------------+------------+----------------------------+"
        Mid(xLinha, 55, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    xLinha = "|          |            |            |            |                            |"
    Mid(xLinha, 2, 10) = Format(pData, "dd/mm/yyyy")
    i = Len(Format(xTotalBruto, "#,###,##0.00"))
    Mid(xLinha, 13 + 12 - i, i) = Format(xTotalBruto, "#,###,##0.00")
    i = Len(Format(xTotalADM, "#,###,##0.00"))
    Mid(xLinha, 26 + 12 - i, i) = Format(xTotalADM, "#,###,##0.00")
    i = Len(Format(xTotalLiquido, "#,###,##0.00"))
    Mid(xLinha, 39 + 12 - i, i) = Format(xTotalLiquido, "#,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
    If chkLinhaSeparadora.Value = 1 Then
        xLinha = "+----------+------------+------------+------------+----------------------------+"
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
    End If
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    If chkLinhaSeparadora.Value = 0 Then
        xLinha = "+----------+------------+------------+------------+----------------------------+"
        BioImprime "@Printer.Print " & xLinha
    End If
    xLinha = "| ** TOTAL |            |            |            |                            |"
    i = Len(Format(lTotalBruto, "#,###,##0.00"))
    Mid(xLinha, 13 + 12 - i, i) = Format(lTotalBruto, "#,###,##0.00")
    i = Len(Format(lTotalADM, "#,###,##0.00"))
    Mid(xLinha, 26 + 12 - i, i) = Format(lTotalADM, "#,###,##0.00")
    i = Len(Format(lTotalLiquido, "#,###,##0.00"))
    Mid(xLinha, 39 + 12 - i, i) = Format(lTotalLiquido, "#,###,##0.00")
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+----------+------------+------------+------------+----------------------------+"
    Mid(xLinha, 55, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
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
Private Sub Relatorio()
    Dim x_data As Date
    Dim Imprimiu As Boolean
    Dim Tempo As Date
    ''Tempo = Time
    
    Imprimiu = False
    ZeraVariaveis
    x_data = CDate(msk_data_i.Text)
    'Loop data
    Do Until x_data > CDate(msk_data_f.Text)
        Call ImpDet(x_data)
        Imprimiu = True
        x_data = x_data + 1
    Loop
    If Imprimiu = True Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório do Resumo de Venda de Cartão de Crédito|@|"
        ''MsgBox "Tempo gasto: " & DateDiff("s", Tempo, Time)
        frm_preview.Show 1
    End If
End Sub
Private Sub cboCartaoCredito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub cboCartaoCredito_LostFocus()
    If cboCartaoCredito.ListIndex <> -1 Then
        lCodigoCartao = Val(Mid(cboCartaoCredito.Text, 1, 2))
        If lCodigoCartao > 0 Then
            If Not CartaoCredito.LocalizarCodigo(lCodigoCartao) Then
                MsgBox "Cartão de Crédito inexistente!", vbInformation, "Erro de Verificação!"
                cboCartaoCredito.SetFocus
                Exit Sub
            End If
        End If
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cboCartaoCredito.SetFocus
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
    cboCartaoCredito.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cboCartaoCredito.SetFocus
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
            Call GravaAuditoria(1, Me.name, 7, "Ref:" & msk_data_i.Text & " a " & msk_data_f.Text & " Comb:" & cboCartaoCredito.Text)
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
    ElseIf cboCartaoCredito.ListIndex = -1 Then
        MsgBox "Selecione um cartão de crédito.", vbInformation, "Atenção!"
        cboCartaoCredito.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lTotalBruto = 0
    lTotalLiquido = 0
    lTotalADM = 0
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    lLocal = 0
    If ValidaCampos Then
        AtivaBotoes (False)
        If SelecionaImpressoraEpson(Me) Then
            DoEvents
            Call GravaAuditoria(1, Me.name, 6, "Ref:" & msk_data_i.Text & " a " & msk_data_f.Text & " Comb:" & cboCartaoCredito.Text)
            Relatorio
        End If
        AtivaBotoes (True)
        cmd_sair.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
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
    PreencheCboCartao
End Sub
Private Sub PreencheCboCartao()
    Dim rstCartaoCredito As New adodb.Recordset
    
    cboCartaoCredito.Clear
    cboCartaoCredito.AddItem "01 - Todos os Cartões"
    Set rstCartaoCredito = Conectar.RsConexao("SELECT Codigo, Nome FROM Cartao_Credito ORDER BY Nome")
    'loop RecordSet
    With rstCartaoCredito
        If Not .BOF Or Not .EOF Then
            .MoveFirst
            Do Until .EOF
                cboCartaoCredito.AddItem Format(!Codigo, "00") & " - " & !Nome
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rstCartaoCredito = Nothing
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
        cboCartaoCredito.SetFocus
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
Private Sub optEmissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub optVencimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub

