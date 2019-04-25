VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form EmissaoLivroPrecoDiferenciado 
   Caption         =   "Emissão do Livro de Preço Diferenciado"
   ClientHeight    =   2745
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "EmissaoLivroPrecoDiferenciado.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "EmissaoLivroPrecoDiferenciado.frx":030A
   ScaleHeight     =   2745
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "EmissaoLivroPrecoDiferenciado.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Visualiza livro de preço diferenciado."
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "EmissaoLivroPrecoDiferenciado.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprime livro de preço diferenciado."
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "EmissaoLivroPrecoDiferenciado.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   11
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
      Begin VB.CheckBox chkExclusivoLoja 
         Caption         =   "Exclusivo da Loja"
         Height          =   255
         Left            =   3660
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkExclusivoPosto 
         Caption         =   "Exclusivo do Posto"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   1875
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "EmissaoLivroPrecoDiferenciado.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbo_grupo 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   4755
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
         Caption         =   "I&mprimir Produto"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Grupo"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "&Data de Emissão"
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
Attribute VB_Name = "EmissaoLivroPrecoDiferenciado"
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
Dim lQtdMaximaLinha As Integer
'Fim de variáveis padrão para relatório
Dim lNomeEmpresa(0 To 50) As String


Dim lSQL As String
Private rsTabela As New adodb.Recordset
Private rsTabela2 As New adodb.Recordset
Private rsTabela3 As New adodb.Recordset

Private Estoque As New cEstoque
Private Produto As New cProduto

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
    Set Estoque = Nothing
    Set Produto = Nothing
End Sub
Private Sub PreencheCboGrupo()
    cbo_grupo.Clear
    cbo_grupo.AddItem "Todos os Grupos"
    cbo_grupo.ItemData(cbo_grupo.NewIndex) = 0
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Nome, Codigo"
    lSQL = lSQL & "  FROM Grupo"
    lSQL = lSQL & " ORDER BY Nome"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            cbo_grupo.AddItem rsTabela("Nome").Value
            cbo_grupo.ItemData(cbo_grupo.NewIndex) = rsTabela("Codigo").Value
            rsTabela.MoveNext
        Loop
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
End Sub
Private Sub AtribuiNomeEmpresas()
    Dim i As Integer
    
    For i = 0 To 50
        lNomeEmpresa(i) = ""
    Next
    
    lSQL = ""
    lSQL = lSQL & "SELECT Codigo, Nome"
    lSQL = lSQL & "  FROM Empresas"
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            lNomeEmpresa(rsTabela("Codigo").Value) = rsTabela("Nome").Value
            rsTabela.MoveNext
        Loop
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    If g_impressora_matricial Then
        lQtdMaximaLinha = 60
    Else
        lQtdMaximaLinha = 95
    End If
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Produto.Nome, Produto.[Codigo do Grupo], Produto.Codigo, Produto.Unidade"
    lSQL = lSQL & "  FROM Produto"
    lSQL = lSQL & " WHERE Produto.Inativo = " & preparaBooleano(False)
    If cbo_grupo.ItemData(cbo_grupo.ListIndex) > 0 Then
        lSQL = lSQL & "   AND Produto.[Codigo do Grupo] = " & cbo_grupo.ItemData(cbo_grupo.ListIndex)
    End If
    If chkExclusivoPosto.Value = 1 And chkExclusivoLoja.Value = 0 Then
        lSQL = lSQL & "   AND Produto.[Exclusivo Posto] = " & preparaBooleano(True)
    End If
    If chkExclusivoLoja.Value = 1 And chkExclusivoPosto.Value = 0 Then
        lSQL = lSQL & "   AND Produto.[Exclusivo Loja] = " & preparaBooleano(True)
    End If
    lSQL = lSQL & " ORDER BY Produto.Nome"
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
End Sub
Private Sub ImpDados()
    LoopTabelaProduto
    If lPagina > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório do Livro de Preço Diferenciado|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub LoopTabelaProduto()
    'loop tabela produto
    Dim xLinha As String
    
    Do Until rsTabela.EOF
        If lPagina = 0 Then
            ImpCab
        End If
        If lLinha >= lQtdMaximaLinha Then
            xLinha = "+-----------+----------------------------------------------+---------+-----------------+-----------------------------------------+------+"
            Mid(xLinha, 17, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & xLinha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        'Verifica se produto tem preço diferente nas empresas
        lSQL = ""
        lSQL = lSQL & "SELECT [Preco de Venda], COUNT(1) AS Quantidade"
        lSQL = lSQL & "  FROM Estoque"
        lSQL = lSQL & " WHERE [Codigo do Produto2] = " & rsTabela("Codigo").Value
        lSQL = lSQL & " GROUP BY [Preco de venda]"
        lSQL = lSQL & " ORDER BY [preco de venda] DESC"
        Set rsTabela2 = New adodb.Recordset
        Set rsTabela2 = Conectar.RsConexao(lSQL)
        If rsTabela2.RecordCount > 1 Then
            xLinha = "+-----------+----------------------------------------------+---------+-----------------+-----------------------------------------+------+"
            BioImprime "@Printer.Print " & xLinha
            lLinha = lLinha + 1
            Do Until rsTabela2.EOF
                'Busca o código da empresa de um preço específico
                lSQL = ""
                lSQL = lSQL & "SELECT Empresa"
                lSQL = lSQL & "  FROM Estoque"
                lSQL = lSQL & " WHERE [Codigo do Produto2] = " & rsTabela("Codigo").Value
                lSQL = lSQL & "   AND [Preco de Venda] = " & preparaValor(rsTabela2("Preco de Venda").Value)
                Set rsTabela3 = New adodb.Recordset
                Set rsTabela3 = Conectar.RsConexao(lSQL)
                If rsTabela3.RecordCount > 0 Then
                    Call ImpDet(rsTabela("Codigo").Value, rsTabela("Nome").Value, rsTabela("unidade").Value, rsTabela2("Preco de Venda").Value, lNomeEmpresa(rsTabela3("Empresa").Value), Val(rsTabela3.RecordCount))
                End If
                If rsTabela3.State = 1 Then
                    rsTabela3.Close
                End If
                rsTabela2.MoveNext
            Loop
        End If
        If rsTabela2.State = 1 Then
            rsTabela2.Close
        End If
        rsTabela.MoveNext
    Loop
End Sub
Private Sub ImpDet(ByVal pCodigo As Long, ByVal pNome As String, ByVal pUnidade As String, ByVal pPrecoVenda As Currency, ByVal pNomeEmpresa As String, ByVal pQtdEmpresa As Integer)
    Dim xLinha As String
    Dim i As Integer
    
    xLinha = "|           |                                              |         |                 |                    |                    |      |"
    i = Len(Format(pCodigo, "#,000"))
    Mid(xLinha, 5 + 5 - i, i) = Format(pCodigo, "#,000")
    Mid(xLinha, 17, 40) = pNome
    Mid(xLinha, vbInformation, 3) = pUnidade
    i = Len(Format(pPrecoVenda, "###,##0.00"))
    Mid(xLinha, 74 + 10 - i, i) = Format(pPrecoVenda, "###,###,##0.00")
    Mid(xLinha, 90, 40) = pNomeEmpresa
    i = Len(Format(pQtdEmpresa, "#0"))
    Mid(xLinha, 133 + 2 - i, i) = Format(pQtdEmpresa, "#0")
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    xLinha = "+-----------+----------------------------------------------+---------+-----------------+-----------------------------------------+------+"
    Mid(xLinha, 17, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
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
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.CurrentY = 0"
    xLinha = "+---------------------------------------------------------------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = True"
    xLinha = "|                                                                                                                           Página: ___ |"
    Mid(xLinha, 3, 40) = g_nome_empresa
    Mid(xLinha, 133, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    '                        1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '               12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    '                                                                                                            123456789012345678901234567890
    xLinha = "| PREÇO DE VENDA DIFERENCIADO                              GRUPO.:                                                   CIDADE, __/__/____ |"
    Mid(xLinha, 68, 30) = cbo_grupo.Text
    i = Len(g_cidade_empresa)
    Mid(xLinha, 94 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 126, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    xLinha = "+-----------+----------------------------------------------+---------+-----------------+-----------------------------------------+------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|  CODIGO   |   DISCRIMINAÇÃO DOS PRODUTOS                 | UNIDADE |  PREÇO DE VENDA | EMPRESA                                 | QTD. |"
    BioImprime "@Printer.Print " & xLinha
'    If chk_linha_separadora.Value = 0 Then
'        xLinha = "+-----------+----------------------------------------------+---------+-----------------+--------------------+--------------------+------+"
'        BioImprime "@Printer.Print " & xLinha
'    End If
End Sub
Private Sub cbo_grupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub chkExclusivoLoja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_grupo.SetFocus
    End If
End Sub
Private Sub chkExclusivoPosto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chkExclusivoLoja.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data.Text = RetiraGString(1)
    chkExclusivoPosto.SetFocus
    g_string = ""
End Sub
Private Sub cmd_imprimir_Click()
    lLocal = 1
    If ValidaCampos Then
        AtivaBotoes (False)
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        End If
        AtivaBotoes (True)
        cmd_sair.SetFocus
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf cbo_grupo.ListIndex = -1 Then
        MsgBox "Escolha o grupo.", vbInformation, "Atenção!"
        cbo_grupo.SetFocus
    ElseIf chkExclusivoPosto.Value = 0 And chkExclusivoLoja.Value = 0 Then
        MsgBox "Selecione se é exclusivo posto e/ou exclusivo loja.", vbInformation, "Atenção!"
        chkExclusivoPosto.SetFocus
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
        AtivaBotoes (False)
        If SelecionaImpressoraEpson(Me) Then
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
        chkExclusivoPosto.Value = 1
        chkExclusivoLoja.Value = 0
        cbo_grupo.ListIndex = 0
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
    
    PreencheCboGrupo
    AtribuiNomeEmpresas
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chkExclusivoPosto.SetFocus
    End If
End Sub
