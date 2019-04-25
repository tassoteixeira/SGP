VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form EmissaoMovimentoCaixa 
   Caption         =   "Emissão do Movimento do Caixa"
   ClientHeight    =   3735
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   7530
   Icon            =   "EmissaoMovimentoCaixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "EmissaoMovimentoCaixa.frx":030A
   ScaleHeight     =   3735
   ScaleWidth      =   7530
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1440
      Picture         =   "EmissaoMovimentoCaixa.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Visualiza movimeno do caixa."
      Top             =   2820
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3300
      Picture         =   "EmissaoMovimentoCaixa.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Imprime movimeno do caixa."
      Top             =   2820
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5160
      Picture         =   "EmissaoMovimentoCaixa.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2820
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7275
      Begin VB.ComboBox cboPeriodoI 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   960
         Width           =   495
      End
      Begin VB.ComboBox cboPeriodoF 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   960
         Width           =   495
      End
      Begin VB.ComboBox cboUsuario 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2220
         Width           =   5475
      End
      Begin VB.ComboBox cboTipoCaixa 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1380
         Width           =   5475
      End
      Begin VB.TextBox txtDataFinal 
         Height          =   285
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtDataInicial 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtDataEmissao 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cboLancamentoPadrao 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1800
         Width           =   5475
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "EmissaoMovimentoCaixa.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2820
         Picture         =   "EmissaoMovimentoCaixa.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   6300
         Picture         =   "EmissaoMovimentoCaixa.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "&Período inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1020
         Width           =   1515
      End
      Begin VB.Label Label7 
         Caption         =   "Período &final"
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   1020
         Width           =   1035
      End
      Begin VB.Label Label6 
         Caption         =   "&Usuário"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   2220
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Tipo de Caixa"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   1380
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "&Tipo de Movimento"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   1800
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
         Left            =   4080
         TabIndex        =   7
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "EmissaoMovimentoCaixa"
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
Dim lTotal As Currency
Dim lSQL As String
Dim rstMovimentoCaixa As adodb.Recordset
Dim rsTabela As adodb.Recordset

Private Usuario As New cUsuario
Private LancamentoPadrao As New cLancamentoPadrao
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set LancamentoPadrao = Nothing
    Set Usuario = Nothing
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lTotal = 0
End Sub
Private Sub PreencheCboUsuario()
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Codigo, Nome"
    lSQL = lSQL & "     FROM Usuario"
    lSQL = lSQL & " ORDER BY Nome"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    
    cboUsuario.Clear
    cboUsuario.AddItem "Todos os Usuarios"
    cboUsuario.ItemData(cboUsuario.NewIndex) = 0
    If rsTabela.RecordCount > 0 Then
        rsTabela.MoveFirst
        Do Until rsTabela.EOF
            cboUsuario.AddItem rsTabela("Nome").Value
            cboUsuario.ItemData(cboUsuario.NewIndex) = rsTabela("Codigo").Value
            rsTabela.MoveNext
        Loop
    End If
    rsTabela.Close
    Set rsTabela = Nothing
End Sub
Private Sub PreencheCboLancamentoPadrao()
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Codigo, Nome"
    lSQL = lSQL & "     FROM LancamentoPadrao"
    lSQL = lSQL & "    WHERE Empresa = " & g_empresa
    lSQL = lSQL & " ORDER BY Nome"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    
    cboLancamentoPadrao.Clear
    cboLancamentoPadrao.AddItem "Todos os Movimentos"
    cboLancamentoPadrao.ItemData(cboLancamentoPadrao.NewIndex) = 0
    If rsTabela.RecordCount > 0 Then
        rsTabela.MoveFirst
        Do Until rsTabela.EOF
            cboLancamentoPadrao.AddItem rsTabela("Nome").Value
            cboLancamentoPadrao.ItemData(cboLancamentoPadrao.NewIndex) = rsTabela("Codigo").Value
            rsTabela.MoveNext
        Loop
    End If
    rsTabela.Close
    Set rsTabela = Nothing
End Sub
Private Sub PreencheCboPeriodos()
    Dim i As Integer
    
    cboPeriodoI.Clear
    cboPeriodoF.Clear
    For i = 1 To 4
        cboPeriodoI.AddItem i
        cboPeriodoF.AddItem i
        cboPeriodoF.ItemData(cboPeriodoF.NewIndex) = i
        cboPeriodoI.ItemData(cboPeriodoI.NewIndex) = i
    Next
End Sub
Private Sub PreencheCboTipoCaixa()
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Codigo, Nome"
    lSQL = lSQL & "     FROM TipoMovimentoCaixa"
    lSQL = lSQL & " ORDER BY Nome"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    
    cboTipoCaixa.Clear
    cboTipoCaixa.AddItem "Todos os Caixas"
    cboTipoCaixa.ItemData(cboTipoCaixa.NewIndex) = 0
    If rsTabela.RecordCount > 0 Then
        rsTabela.MoveFirst
        Do Until rsTabela.EOF
            cboTipoCaixa.AddItem rsTabela("Nome").Value
            cboTipoCaixa.ItemData(cboTipoCaixa.NewIndex) = rsTabela("Codigo").Value
            rsTabela.MoveNext
        Loop
    End If
    rsTabela.Close
    Set rsTabela = Nothing
End Sub
Private Sub Relatorio()
    
    ZeraVariaveis
    
    lSQL = ""
    lSQL = lSQL & " SELECT Data, [Numero do Documento], Periodo, Complemento, Valor, [Codigo do Usuario], [Codigo do Lancamento Padrao]"
    lSQL = lSQL & "   FROM MovimentoCaixaPista"
    lSQL = lSQL & "  WHERE Empresa = " & g_empresa
    lSQL = lSQL & "    AND Data >= " & preparaData(CDate(txtDataInicial.Text))
    lSQL = lSQL & "    AND Data <= " & preparaData(CDate(txtDataFinal.Text))
    lSQL = lSQL & "    AND Periodo >= " & Val(cboPeriodoI.Text)
    lSQL = lSQL & "    AND Periodo <= " & Val(cboPeriodoF.Text)
    If cboTipoCaixa.ItemData(cboTipoCaixa.ListIndex) > 0 Then
        lSQL = lSQL & "    AND [Tipo do Movimento] = " & cboTipoCaixa.ItemData(cboTipoCaixa.ListIndex)
    End If
    If cboLancamentoPadrao.ItemData(cboLancamentoPadrao.ListIndex) > 0 Then
        lSQL = lSQL & "    AND [Codigo do Lancamento Padrao] = " & cboLancamentoPadrao.ItemData(cboLancamentoPadrao.ListIndex)
    End If
    If cboUsuario.ItemData(cboUsuario.ListIndex) > 0 Then
        lSQL = lSQL & "    AND [Codigo do Usuario] = " & cboUsuario.ItemData(cboUsuario.ListIndex)
    End If
    lSQL = lSQL & "  ORDER BY Data, Periodo, [Numero do Documento], Complemento"
    Set rstMovimentoCaixa = Conectar.RsConexao(lSQL)
    If rstMovimentoCaixa.RecordCount > 0 Then
        ImpDados
    Else
        MsgBox "Não tem movimento nas condições informadas!", vbInformation, "Relatório não será impresso!"
    End If
    rstMovimentoCaixa.Close
    Set rstMovimentoCaixa = Nothing
   
    Call GravaAuditoria(1, Me.name, 7, "Tipo Cx:" & cboTipoCaixa.ItemData(cboTipoCaixa.ListIndex) & " Ref:" & txtDataInicial.Text & " a " & txtDataFinal.Text)
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    Dim xLinha As String
    'loop movimento de notas de abastecimento
    With rstMovimentoCaixa
        Do Until .EOF
            If lPagina = 0 Then
                ImpCab
                'ImpCliente
            End If
            If lLinha >= 60 Then
                xLinha = "+------------+--------+------+------------------------------------------+----------------+----------------+----------------+------------+"
                BioImprime "@Printer.Print " & xLinha
                BioImprime "@@Printer.NewPage"
                ImpCab
            End If
            ImpDet (False)
            .MoveNext
        Loop
    End With
    If lTotal > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório do Movimento de Caixa|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub ImpDet(ByVal pBaixada As Boolean)
    Dim xLinha As String
    Dim i As Integer
    Dim xValor As Currency
    
    xLinha = Space(137)
    xLinha = "         1         2         3         4         5         6         7         8         9        10        11        12        13     13"
    xLinha = "12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567"
    xLinha = "|12/99/9999| 1212.123 | 1234 | 1234567890123456789012345678901234567890 |   quantidade   |   quantidade   |   quantidade   | empresa    |"
    xLinha = "|          |          |   |                                          |                |                |                   |            |"
    With rstMovimentoCaixa
        lTotal = lTotal + !valor
        Mid(xLinha, 2, 10) = Format(!Data, "dd/mm/yyyy")
        Mid(xLinha, 14, 10) = ![Numero do Documento]
        Mid(xLinha, 25, 1) = !Periodo
        Mid(xLinha, 29, 40) = !Complemento
        i = Len(Format(!valor, "###,###,##0.00"))
        Mid(xLinha, 72 + 14 - i, i) = Format(!valor, "###,###,##0.00")
        If LancamentoPadrao.LocalizarCodigo(g_empresa, ![Codigo do Lancamento Padrao]) Then
            Mid(xLinha, 105, 19) = LancamentoPadrao.Nome
        End If
        If Usuario.LocalizarCodigo(![Codigo do Usuario]) Then
            Mid(xLinha, 125, 12) = Usuario.Nome
        End If
    End With
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    
    xLinha = "+----------+----------+---+------------------------------------------+----------------+----------------+-------------------+------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                                                             Total  |                |                                                 |"
    i = Len(Format(lTotal, "###,###,##0.00"))
    Mid(xLinha, 72 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+--------------------------------------------------------------------+----------------+-------------------------------------------------+"
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
    xLinha = "|                                                                  Página, ___ |"
    Mid(xLinha, 3, 40) = x_string_40
    Mid(xLinha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & xLinha
    '                   1         2         3         4         5         6         7         8
    '          12345678901234567890123456789012345678901234567890123456789012345678901234567890
    '                                              123456789012345678901234567890
    xLinha = "| RELAÇÃO DA MOVIMENTACAO DO CAIXA                          CIDADE, __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 69, 10) = txtDataEmissao.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| DATA DO MOVIMENTO.: __/__/____ A __/__/____     PERIODO..: _ AO _            |"
    Mid(xLinha, 23, 10) = txtDataInicial.Text
    Mid(xLinha, 36, 10) = txtDataFinal.Text
    Mid(xLinha, 62, 1) = cboPeriodoI.Text
    Mid(xLinha, 67, 1) = cboPeriodoF.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| TIPO DE CAIXA.....:                                                          |"
    Mid(xLinha, 23, 30) = cboTipoCaixa.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| TIPO DE MOVIMENTO.:                                                          |"
    Mid(xLinha, 23, 30) = cboLancamentoPadrao.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| USUARIO...........:                                                          |"
    Mid(xLinha, 23, 30) = cboUsuario.Text
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    xLinha = "+----------+----------+---+------------------------------------------+----------------+----------------+-------------------+------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| DATA  DO |  NUMERO  |PER|COMPLEMENTO                               |      VALOR     |                |                   | USUARIO    |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| MOVIMENTO| DOCUMENTO|   |                                          |                |                |                   |            |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+----------+----------+---+------------------------------------------+----------------+----------------+-------------------+------------+"
    BioImprime "@Printer.Print " & xLinha
End Sub
Private Sub cboUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub cboLancamentoPadrao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboUsuario.SetFocus
    End If
End Sub
Private Sub cboPeriodoF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboTipoCaixa.SetFocus
    End If
End Sub
Private Sub cboPeriodoI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboPeriodoF.ListIndex = cboPeriodoI.ListIndex
        cboPeriodoF.SetFocus
    End If
End Sub
Private Sub cboTipoCaixa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboLancamentoPadrao.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = txtDataEmissao.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        txtDataInicial.Text = RetiraGString(1)
        txtDataFinal.Text = RetiraGString(2)
        cboPeriodoI.SetFocus
    Else
        txtDataEmissao.Text = RetiraGString(1)
        txtDataInicial.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_data_f_Click()
    g_string = txtDataFinal.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        txtDataInicial.Text = RetiraGString(1)
        txtDataFinal.Text = RetiraGString(2)
    Else
        txtDataFinal.Text = RetiraGString(1)
    End If
    g_string = ""
    cboPeriodoI.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = txtDataInicial.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        txtDataInicial.Text = RetiraGString(1)
        txtDataFinal.Text = RetiraGString(2)
        cboPeriodoI.SetFocus
    Else
        txtDataInicial.Text = RetiraGString(1)
        txtDataFinal.SetFocus
    End If
    g_string = ""
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
    If Not IsDate(txtDataEmissao.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        txtDataEmissao.SetFocus
    ElseIf Not IsDate(txtDataInicial.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        txtDataInicial.SetFocus
    ElseIf Not IsDate(txtDataFinal.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        txtDataFinal.SetFocus
    ElseIf CDate(txtDataFinal.Text) < CDate(txtDataInicial.Text) Then
        MsgBox "Data final deve ser maior que a data inicial.", vbInformation, "Atenção!"
        txtDataFinal.SetFocus
    ElseIf cboLancamentoPadrao.ListIndex = -1 Then
        MsgBox "Selecione um grupo.", vbInformation, "Atenção!"
        cboLancamentoPadrao.SetFocus
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
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(txtDataEmissao.Text) Then
        txtDataEmissao.Text = Format(g_data_def, "dd/mm/yyyy")
        txtDataInicial.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        txtDataFinal.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        cboPeriodoI.ListIndex = 0
        cboPeriodoF.ListIndex = 0
        cboTipoCaixa.ListIndex = 0
        cboLancamentoPadrao.ListIndex = 0
        cboUsuario.ListIndex = 0
        txtDataInicial.SetFocus
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
    PreencheCboPeriodos
    PreencheCboTipoCaixa
    PreencheCboLancamentoPadrao
    PreencheCboUsuario
End Sub
Private Sub txtDataEmissao_GotFocus()
    txtDataEmissao.Text = fDesmascaraData(txtDataEmissao.Text)
    txtDataEmissao.SelStart = 0
    txtDataEmissao.SelLength = 4
    txtDataEmissao.MaxLength = 8
End Sub
Private Sub txtDataEmissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtDataInicial.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtDataEmissao_LostFocus()
    txtDataEmissao.MaxLength = 10
    txtDataEmissao.Text = fMascaraData(txtDataEmissao.Text)
End Sub
Private Sub txtDataFinal_GotFocus()
    txtDataFinal.Text = fDesmascaraData(txtDataFinal.Text)
    txtDataFinal.SelStart = 0
    txtDataFinal.SelLength = 4
    txtDataFinal.MaxLength = 8
End Sub
Private Sub txtDataFinal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboPeriodoI.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtDataFinal_LostFocus()
    txtDataFinal.MaxLength = 10
    txtDataFinal.Text = fMascaraData(txtDataFinal.Text)
End Sub
Private Sub txtDataInicial_GotFocus()
    txtDataInicial.Text = fDesmascaraData(txtDataInicial.Text)
    txtDataInicial.SelStart = 0
    txtDataInicial.SelLength = 4
    txtDataInicial.MaxLength = 8
End Sub
Private Sub txtDataInicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtDataFinal.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtDataInicial_LostFocus()
    txtDataInicial.MaxLength = 10
    txtDataInicial.Text = fMascaraData(txtDataInicial.Text)
    If IsDate(txtDataInicial.Text) Then
        txtDataFinal.Text = txtDataInicial.Text
    End If
End Sub
