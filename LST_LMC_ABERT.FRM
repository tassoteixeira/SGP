VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_lmc_abertura 
   Caption         =   "Emissão do Termo de Abertura do L. M. C."
   ClientHeight    =   4665
   ClientLeft      =   2985
   ClientTop       =   2115
   ClientWidth     =   6975
   Icon            =   "LST_LMC_ABERT.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "LST_LMC_ABERT.frx":030A
   ScaleHeight     =   4665
   ScaleWidth      =   6975
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   1260
      Picture         =   "LST_LMC_ABERT.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   3720
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3120
      Picture         =   "LST_LMC_ABERT.frx":17C2
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Imprime o termo de abertura do L.M.C."
      Top             =   3720
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4980
      Picture         =   "LST_LMC_ABERT.frx":2DCC
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3720
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.CheckBox chkImprimeMoldura 
         Caption         =   "Imprime moldura nos termos de Abertura/Fechamento"
         Height          =   195
         Left            =   1680
         TabIndex        =   16
         Top             =   3120
         Value           =   1  'Checked
         Width           =   4875
      End
      Begin VB.CheckBox chkNumeraTermos 
         Caption         =   "Imprime número de paginas nos termos"
         Height          =   195
         Left            =   1680
         TabIndex        =   15
         Top             =   2820
         Value           =   1  'Checked
         Width           =   4875
      End
      Begin VB.TextBox txtPaginaInicial 
         Height          =   300
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   12
         Top             =   1980
         Width           =   675
      End
      Begin VB.TextBox txtPaginaFinal 
         Height          =   300
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   14
         Top             =   2400
         Width           =   675
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2700
         Picture         =   "LST_LMC_ABERT.frx":445E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   300
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2700
         Picture         =   "LST_LMC_ABERT.frx":5738
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   720
         Width           =   495
      End
      Begin VB.ComboBox cbo_combustivel 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1560
         Width           =   4935
      End
      Begin VB.TextBox txt_numero_livro 
         Height          =   300
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   8
         Top             =   1140
         Width           =   675
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   720
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
         TabIndex        =   2
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Caption         =   "&Combustível"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "D&ata Final"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Livro de Número"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Página Inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   1980
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Página Final"
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "&Data Inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   180
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_lmc_abertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_local As Integer
Dim l_linha_tab As Currency
Dim l_coluna_tab As Currency
Dim l_pagina As Integer
Dim lTipoCombustivel As String
Dim lTipoLivro As String
Dim lDataInicial As Date
Dim lChamaPesquisa As Boolean

Private Empresa As New cEmpresa
Private LivroLMC As New cLivroLMC
Private TanqueCombustivel As New cTanqueCombustivel

Private rsCombustivel As New adodb.Recordset
Private rsTanque As New adodb.Recordset
Private rsBico As New adodb.Recordset
Private Sub AtivaBotoes(ByVal pAtiva As Boolean)
    'cmd_visualizar.Enabled = pAtiva
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
Private Sub AtualTela()
    Dim i As Integer
    
    If IsDate(lDataInicial) Then
        If LivroLMC.LocalizarCodigo(g_empresa, lTipoCombustivel, lTipoLivro, lDataInicial) Then
            msk_data_i.Text = LivroLMC.DataInicial
            msk_data_f.Text = LivroLMC.DataFinal
            txt_numero_livro.Text = LivroLMC.NumeroLivro
            cbo_combustivel.ListIndex = -1
            For i = 0 To cbo_combustivel.ListCount - 1
                cbo_combustivel.ListIndex = i
                If LivroLMC.TipoCombustivel = "TC" And Mid(cbo_combustivel.Text, 1, 2) = "**" Then
                    Exit For
                End If
                If Mid(cbo_combustivel.Text, 1, 2) = LivroLMC.TipoCombustivel Then
                    Exit For
                End If
            Next
            txtPaginaInicial.Text = Format(LivroLMC.PaginaInicial, "000")
            txtPaginaFinal.Text = Format(LivroLMC.PaginaFinal, "000")
        Else
            MsgBox "Não foi possível localizar registro", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Empresa = Nothing
    Set LivroLMC = Nothing
    Set TanqueCombustivel = Nothing
End Sub
Private Sub PesquisaLivro()
    g_string = "T"
    ConsultaLivroLMC.Show 1
    If Len(g_string) > 0 Then
        lTipoLivro = RetiraGString(1)
        lTipoCombustivel = RetiraGString(2)
        lDataInicial = RetiraGString(3)
        g_string = ""
        AtualTela
    End If
    lChamaPesquisa = False
    cmd_imprimir.SetFocus
End Sub
Private Sub PreencheCboCombustivel()
    Dim xSQL As String
    
    cbo_combustivel.Clear
    cbo_combustivel.AddItem "** - Todos Os Combustíveis"
    
    'Prepara SQL
    xSQL = ""
    xSQL = xSQL & "SELECT Codigo, Nome, Ordem"
    xSQL = xSQL & "  FROM Combustivel"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & " ORDER BY Ordem, Nome"
    'Abre RecordSet
    Set rsCombustivel = New adodb.Recordset
    Set rsCombustivel = Conectar.RsConexao(xSQL)
    'Loop no RecordSet
    If rsCombustivel.RecordCount > 0 Then
        rsCombustivel.MoveFirst
        Do Until rsCombustivel.EOF
            cbo_combustivel.AddItem rsCombustivel("Codigo").Value & " - " & rsCombustivel("Nome").Value
            rsCombustivel.MoveNext
        Loop
    End If
End Sub
Private Function PreparaTanque(ByVal pCombustivel As String) As String
    Dim xSQL As String
    Dim i As Integer
    
    PreparaTanque = ""
    i = 0
    'Prepara SQL
    xSQL = ""
    xSQL = xSQL & "SELECT [Numero do Tanque]"
    xSQL = xSQL & "  FROM Tanque_Combustivel"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(pCombustivel)
    xSQL = xSQL & " ORDER BY [Numero do Tanque]"
    'Abre RecordSet
    Set rsTanque = New adodb.Recordset
    Set rsTanque = Conectar.RsConexao(xSQL)

    'Loop no RecordSet
    If rsTanque.RecordCount > 0 Then
        rsTanque.MoveFirst
        Do Until rsTanque.EOF
            i = i + 1
            If i > 1 Then
                PreparaTanque = PreparaTanque & ", "
            End If
            PreparaTanque = PreparaTanque & rsTanque("Numero do Tanque").Value
            rsTanque.MoveNext
        Loop
    End If
    rsTanque.Close
    Set rsTanque = Nothing
End Function
Private Function PreparaBico(ByVal pCombustivel As String) As String
    Dim xSQL As String
    Dim i As Integer
    
    PreparaBico = ""
    i = 0
    'Prepara SQL
    xSQL = ""
    xSQL = xSQL & "SELECT Codigo"
    xSQL = xSQL & "  FROM Bomba"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(pCombustivel)
    xSQL = xSQL & " ORDER BY Codigo"
    'Abre RecordSet
    Set rsBico = New adodb.Recordset
    Set rsBico = Conectar.RsConexao(xSQL)

    'Loop no RecordSet
    If rsBico.RecordCount > 0 Then
        rsBico.MoveFirst
        Do Until rsBico.EOF
            i = i + 1
            If i > 1 Then
                PreparaBico = PreparaBico & ", "
            End If
            PreparaBico = PreparaBico & rsBico("Codigo").Value
            rsBico.MoveNext
        Loop
    End If
    rsBico.Close
    Set rsBico = Nothing
End Function
Private Sub Relatorio()
    ZeraVariaveis
    ImpDados (True)
    ImpDados (False)
End Sub
Private Sub ImpDados(ByVal pAbertura As Boolean)
Dim xLinha As Currency
Dim x_mes As String
Dim x_data As Date
Dim xString As String
Dim xQtdFolha As Integer

    ImpGrade (pAbertura)
    If pAbertura = False Then
        x_data = msk_data_f.Text
    Else
        x_data = msk_data_i.Text
    End If
    x_mes = Format(x_data, "mmmm")
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    Printer.FontName = "Times New Roman"
    Printer.FontName = "Times New Roman"
    Printer.FontSize = 12
    Printer.FontBold = True
    ImprimeTexto txt_numero_livro.Text, l_coluna_tab + 9.4, l_coluna_tab + 18, l_linha_tab + 5 - 0.13, l_local
    xQtdFolha = Val(txtPaginaFinal.Text) - Val(txtPaginaInicial.Text) + 1
    ImprimeTexto Format(xQtdFolha, "##000"), l_coluna_tab + 7.2, l_coluna_tab + 18, l_linha_tab + 7.5 - 0.13, l_local
    ImprimeTexto txtPaginaInicial.Text, l_coluna_tab + 14, l_coluna_tab + 18, l_linha_tab + 7.5 - 0.13, l_local
    ImprimeTexto msk_data_i.Text, l_coluna_tab + 4.5, l_coluna_tab + 7.5, l_linha_tab + 9.3, l_local
    ImprimeTexto msk_data_f.Text, l_coluna_tab + 8.5, l_coluna_tab + 11.5, l_linha_tab + 9.3, l_local
    Printer.FontSize = 14
    If Mid(cbo_combustivel.Text, 1, 2) = "**" Then
        xLinha = 2
        rsCombustivel.MoveFirst
        Do Until rsCombustivel.EOF
            xLinha = xLinha + 0.5
            ImprimeCentralizado rsCombustivel("Nome").Value, l_coluna_tab + 0, l_coluna_tab + 18, l_linha_tab + xLinha, l_local
            rsCombustivel.MoveNext
        Loop
    Else
        ImprimeCentralizado Mid(cbo_combustivel.Text, 6, Len(cbo_combustivel.Text)), l_coluna_tab + 0, l_coluna_tab + 18, l_linha_tab + 2.5, l_local
    End If
    Printer.FontSize = 12
    Printer.FontBold = True
    ImprimeTexto txtPaginaFinal.Text, l_coluna_tab + 2.5, l_coluna_tab + 18, l_linha_tab + 8.1 - 0.13, l_local
    ImprimeTexto g_nome_empresa, l_coluna_tab + 4, l_coluna_tab + 18, l_linha_tab + 10 - 0.13, l_local
    ImprimeTexto Trim(Empresa.Endereco) & " - " & Empresa.Bairro, l_coluna_tab + 4.7, l_coluna_tab + 18, l_linha_tab + 10.5 - 0.13, l_local
    ImprimeTexto Empresa.Cidade, l_coluna_tab + 4.5, l_coluna_tab + 18, l_linha_tab + 11 - 0.13, l_local
    ImprimeTexto Empresa.Estado, l_coluna_tab + 12.1, l_coluna_tab + 18, l_linha_tab + 11 - 0.13, l_local
    ImprimeTexto fMascaraCNPJ(Empresa.CGC), l_coluna_tab + 10.4, l_coluna_tab + 17, l_linha_tab + 11.5 - 0.13, l_local
    ImprimeTexto Empresa.InscricaoEstadual, l_coluna_tab + 5.7, l_coluna_tab + 18, l_linha_tab + 12 - 0.13, l_local
    ImprimeTexto Empresa.Distribuidora, l_coluna_tab + 7, l_coluna_tab + 17.9, l_linha_tab + 13 - 0.13, l_local
    If Mid(cbo_combustivel.Text, 1, 2) = "**" Then
        xLinha = 14 - 0.08 - 0.5
        rsCombustivel.MoveFirst
        Do Until rsCombustivel.EOF
            xLinha = xLinha + 0.5
            xString = Format(TanqueCombustivel.TotalCapacidadeArmazenamento(g_empresa, rsCombustivel("Codigo").Value), "###,###")
            xString = xString & " Lts de " & rsCombustivel("Nome").Value
            ImprimeTexto xString, l_coluna_tab + 8.3, l_coluna_tab + 18, l_linha_tab + xLinha, l_local
            xLinha = xLinha + 0.5
            ImprimeTexto PreparaTanque(rsCombustivel("Codigo").Value), l_coluna_tab + 7, l_coluna_tab + 18, l_linha_tab + xLinha, l_local
            'ImprimeTexto PreparaBico(rsCombustivel("Codigo").Value), l_coluna_tab + 13.5, l_coluna_tab + 18, l_linha_tab + xLinha, l_local
            xLinha = xLinha + 0.5
            ImprimeTexto PreparaBico(rsCombustivel("Codigo").Value), l_coluna_tab + 7.7, l_coluna_tab + 18, l_linha_tab + xLinha, l_local
            rsCombustivel.MoveNext
        Loop
    Else
        ImprimeTexto Format(TanqueCombustivel.TotalCapacidadeArmazenamento(g_empresa, Mid(cbo_combustivel.Text, 1, 2)), "###,###"), l_coluna_tab + 8.3, l_coluna_tab + 18, l_linha_tab + 14 - 0.13, l_local
        ImprimeTexto PreparaTanque(Mid(cbo_combustivel.Text, 1, 2)), l_coluna_tab + 7, l_coluna_tab + 18, l_linha_tab + 14.5, l_local
        'ImprimeTexto PreparaBico(Mid(cbo_combustivel.Text, 1, 2)), l_coluna_tab + 13.5, l_coluna_tab + 18, l_linha_tab + 14.5, l_local
        ImprimeTexto PreparaBico(Mid(cbo_combustivel.Text, 1, 2)), l_coluna_tab + 7.7, l_coluna_tab + 18, l_linha_tab + 15, l_local
    End If
    ImprimeTexto Trim(Empresa.Cidade) & ", " & Day(x_data) & " de " & x_mes & " de " & Year(x_data), l_coluna_tab + 6, l_coluna_tab + 18, l_linha_tab + 20, l_local
    ImprimeCentralizado Trim(Empresa.ResponsavelLegal), l_coluna_tab + 0.5, l_coluna_tab + 8.5, l_linha_tab + 22, l_local
    If g_nome_empresa = "AUTO POSTO BRISA LTDA" Then
        'Responsável pela Empresa
        ImprimeCentralizado "CPF: 810.190.921-49", l_coluna_tab + 0.5, l_coluna_tab + 8.5, l_linha_tab + 22.5, l_local
        ImprimeCentralizado "SÓCIO ADMINISTRADOR", l_coluna_tab + 0.5, l_coluna_tab + 8.5, l_linha_tab + 23, l_local
        
        'Contador
        ImprimeCentralizado "JOSE CARLOS MOREIRA", l_coluna_tab + 9.5, l_coluna_tab + 17.5, l_linha_tab + 22, l_local
        ImprimeCentralizado "CPF: 232.475.981-15", l_coluna_tab + 9.5, l_coluna_tab + 17.5, l_linha_tab + 22.5, l_local
        ImprimeCentralizado "CRC GO: 014170/O-0", l_coluna_tab + 9.5, l_coluna_tab + 17.5, l_linha_tab + 23, l_local
    End If
    
    Printer.FontBold = False
    Printer.EndDoc
End Sub
Private Sub ImpGrade(ByVal pAbertura As Boolean)
Dim i As Integer
Dim x_termo As String
Dim xPagina As String
Dim xLinha As Currency

    l_coluna_tab = 2
    l_linha_tab = 1
    If pAbertura Then
        x_termo = "T E R M O     D E     A B E R T U R A"
        If chkNumeraTermos.Value = 1 Then
            xPagina = "Folha: " & txtPaginaInicial.Text
        End If
    Else
        x_termo = "T E R M O     D E     E N C E R R A M E N T O"
        If chkNumeraTermos.Value = 1 Then
            xPagina = "Folha: " & txtPaginaFinal.Text
        End If
    End If
    
    'Printer.ForeColor = RGB(256, 0, 0) 'dados em vermelho
    Printer.ScaleMode = 7
    Printer.FontName = "Times New Roman"
    Printer.FontName = "Times New Roman"
    Printer.FontSize = 10
    Printer.FontBold = False
    
    If chkImprimeMoldura.Value = 1 Then
        Printer.DrawWidth = 8
        Printer.Line (l_coluna_tab, l_linha_tab)-(l_coluna_tab + 18, l_linha_tab + 26), RGB(0, 0, 0), B
    End If
    Printer.DrawWidth = 2
    
    Printer.FontSize = 14
    Printer.FontBold = True
    ImprimeCentralizado "LIVRO DE MOVIMENTAÇÃO DE COMBUSTÍVEIS (LMC)", l_coluna_tab + 0, l_coluna_tab + 18, l_linha_tab + 1, l_local
    
    Printer.FontSize = 12
    Printer.FontBold = True
    ImprimeTexto xPagina, l_coluna_tab + 15, l_coluna_tab + 18, l_linha_tab + 5.5, l_local
    
    Printer.FontSize = 8
    Printer.FontBold = False
    ImprimeCentralizado "(Portaria N. 26, de 13 de Novembro de 1.992 do Departamento Nacional de Combustíveis)", l_coluna_tab + 0, l_coluna_tab + 18, l_linha_tab + 1.6, l_local
    Printer.FontSize = 10
    ImprimeCentralizado "Livro de n.           ", l_coluna_tab, l_coluna_tab + 18, l_linha_tab + 5, l_local
    Printer.FontSize = 14
    Printer.FontBold = True
    ImprimeCentralizado x_termo, l_coluna_tab + 0, l_coluna_tab + 18, l_linha_tab + 6, l_local
    Printer.FontSize = 10
    Printer.FontBold = False
    ImprimeTexto "Este livro que contém           páginas numeradas eletrônicamente de n.           ao n.", l_coluna_tab + 4, l_coluna_tab + 18, l_linha_tab + 7.5, l_local
    If pAbertura Then
        ImprimeTexto "e  servirá  para  o   lançamento   das   operações   próprias   do   estabelecimento   do", l_coluna_tab + 2.5 + 1, l_coluna_tab + 18, l_linha_tab + 8.1, l_local
    Else
        ImprimeTexto "e  serviu  para  o   lançamento   das   operações   próprias   do   estabelecimento   do", l_coluna_tab + 2.5 + 1, l_coluna_tab + 18, l_linha_tab + 8.1, l_local
    End If
    ImprimeTexto "contribuinte abaixo identificado:", l_coluna_tab + 2.5, l_coluna_tab + 18, l_linha_tab + 8.7, l_local
    ImprimeTexto "Referente a:", l_coluna_tab + 2.5, l_coluna_tab + 4.5, l_linha_tab + 9.3, l_local
    ImprimeCentralizado "a", l_coluna_tab + 7.5, l_coluna_tab + 8.5, l_linha_tab + 9.3, l_local
    ImprimeTexto "Da firma", l_coluna_tab + 2.5, l_coluna_tab + 18, l_linha_tab + 10, l_local
    ImprimeTexto "Estabelecida à", l_coluna_tab + 2.5, l_coluna_tab + 18, l_linha_tab + 10.5, l_local
    ImprimeTexto "Na cidade de", l_coluna_tab + 2.5, l_coluna_tab + 18, l_linha_tab + 11, l_local
    ImprimeTexto "Estado de", l_coluna_tab + 10.5, l_coluna_tab + 18, l_linha_tab + 11, l_local
    ImprimeTexto "Inscrição no Cadastro Geral de Contribuintes(M.F.) n.", l_coluna_tab + 2.5, l_coluna_tab + 18, l_linha_tab + 11.5, l_local
    ImprimeTexto "Inscrição Estadual n.", l_coluna_tab + 2.5, l_coluna_tab + 18, l_linha_tab + 12, l_local
    ImprimeTexto "Inscrição Municipal n.", l_coluna_tab + 2.5, l_coluna_tab + 18, l_linha_tab + 12.5, l_local
    ImprimeTexto "Operando com a distribuidora", l_coluna_tab + 2.5, l_coluna_tab + 18, l_linha_tab + 13, l_local
    If Mid(cbo_combustivel.Text, 1, 2) = "**" Then
        xLinha = 14 - 0.5
        rsCombustivel.MoveFirst
        Do Until rsCombustivel.EOF
            xLinha = xLinha + 0.5
            ImprimeTexto "Capacidade normal de armazenamento", l_coluna_tab + 2.5, l_coluna_tab + 18, l_linha_tab + xLinha, l_local
            xLinha = xLinha + 0.5
            ImprimeTexto "N. do(s) Tanque(s):", l_coluna_tab + 4, l_coluna_tab + 18, l_linha_tab + xLinha, l_local
            'ImprimeTexto "N. do(s) Bico(s):", l_coluna_tab + 11, l_coluna_tab + 18, l_linha_tab + xLinha, l_local
            xLinha = xLinha + 0.5
            ImprimeTexto "N. do(s) Bico(s):", l_coluna_tab + 5, l_coluna_tab + 18, l_linha_tab + xLinha, l_local
            rsCombustivel.MoveNext
        Loop
    Else
        ImprimeTexto "Capacidade normal de armazenamento", l_coluna_tab + 2.5, l_coluna_tab + 18, l_linha_tab + 14, l_local
        ImprimeTexto "N. do(s) Tanque(s):", l_coluna_tab + 4, l_coluna_tab + 18, l_linha_tab + 14.5, l_local
        'ImprimeTexto "N. do(s) Bico(s):", l_coluna_tab + 11, l_coluna_tab + 18, l_linha_tab + 14.5, l_local
        ImprimeTexto "N. do(s) Bico(s):", l_coluna_tab + 5, l_coluna_tab + 18, l_linha_tab + 15, l_local
    End If
    Printer.DrawWidth = 8
'    Printer.Line (l_coluna_tab + 0.5, l_linha_tab + 22.5)-(l_coluna_tab + 8.5, l_linha_tab + 22.5), RGB(0, 0, 0), B
'    Printer.Line (l_coluna_tab + 9.5, l_linha_tab + 22.5)-(l_coluna_tab + 17.5, l_linha_tab + 22.5), RGB(0, 0, 0), B
'    Printer.Line (l_coluna_tab + 5, l_linha_tab + 24.5)-(l_coluna_tab + 13, l_linha_tab + 24.5), RGB(0, 0, 0), B
    Printer.Line (l_coluna_tab + 0.5, l_linha_tab + 21.5)-(l_coluna_tab + 8.5, l_linha_tab + 21.5), RGB(0, 0, 0), B
    Printer.Line (l_coluna_tab + 9.5, l_linha_tab + 21.5)-(l_coluna_tab + 17.5, l_linha_tab + 21.5), RGB(0, 0, 0), B
    Printer.Line (l_coluna_tab + 5, l_linha_tab + 24.5)-(l_coluna_tab + 13, l_linha_tab + 24.5), RGB(0, 0, 0), B
    Printer.DrawWidth = 2
'    ImprimeCentralizado "Assinatura do Representante Legal da Empresa", l_coluna_tab + 0.5, l_coluna_tab + 8.5, l_linha_tab + 23, l_local
'    ImprimeCentralizado "Assinatura do Contabilista", l_coluna_tab + 9.5, l_coluna_tab + 17.5, l_linha_tab + 23, l_local
'    ImprimeCentralizado "Visto da Autoridade Fazendaria", l_coluna_tab + 5, l_coluna_tab + 13, l_linha_tab + 25, l_local
    ImprimeCentralizado "Assinatura do Representante Legal da Empresa", l_coluna_tab + 0.5, l_coluna_tab + 8.5, l_linha_tab + 21.5, l_local
    ImprimeCentralizado "Assinatura do Contabilista", l_coluna_tab + 9.5, l_coluna_tab + 17.5, l_linha_tab + 21.5, l_local
    ImprimeCentralizado "Visto da Autoridade Fazendaria", l_coluna_tab + 5, l_coluna_tab + 13, l_linha_tab + 25, l_local
End Sub
Private Sub ZeraVariaveis()
    l_coluna_tab = 2
    l_linha_tab = 1
End Sub
Private Sub cbo_combustivel_GotFocus()
    SendMessageLong cbo_combustivel.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_combustivel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub cbo_combustivel_LostFocus()
    Dim xData As String
    Dim i As Integer
    xData = "32" & Mid(msk_data_i.Text, 3, 8)
    Do Until IsDate(xData)
        Mid(xData, 1, 2) = Val(Mid(xData, 1, 2)) - 1
    Loop
    i = Val(Mid(xData, 1, 2))
    If chkNumeraTermos.Value = 1 Then
        If Mid(cbo_combustivel.Text, 1, 2) = "**" Then
            txtPaginaInicial.Text = 1
            txtPaginaFinal.Text = i * 3 + 2
        ElseIf Mid(cbo_combustivel.Text, 1, 2) = "A " Then
            txtPaginaInicial.Text = 1
            txtPaginaFinal.Text = i + 2
        ElseIf Mid(cbo_combustivel.Text, 1, 2) = "D " Then
            txtPaginaInicial.Text = i + 2 + 1
            txtPaginaFinal.Text = (i + 2) * 2
        ElseIf Mid(cbo_combustivel.Text, 1, 2) = "G " Then
            txtPaginaInicial.Text = (i + 2) * 2 + 1
            txtPaginaFinal.Text = (i + 2) * 3
        End If
    Else
        txtPaginaInicial.Text = "001"
        txtPaginaFinal.Text = Format(CDate(msk_data_f.Text) - CDate(msk_data_i.Text) + 1, "000")
    End If
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
    txt_numero_livro.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        txt_numero_livro.SetFocus
    Else
        msk_data_i = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    If ValidaCampos Then
        AtivaBotoes (False)
        If SelecionaImpressoraHP(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        End If
        AtivaBotoes (True)
        cmd_sair.SetFocus
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data_i.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f.Text) < CDate(msk_data_i.Text) Then
        MsgBox "Data final deve ser maior que a data inicial.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf Val(txt_numero_livro.Text) = 0 Then
        MsgBox "Informe o número do livro.", vbInformation, "Atenção!"
        txt_numero_livro.SetFocus
    ElseIf cbo_combustivel.ListIndex = -1 Then
        MsgBox "Selecione um combustível.", vbInformation, "Atenção!"
        cbo_combustivel.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    lChamaPesquisa = True
    Call GravaAuditoria(1, Me.name, 5, "")
    PesquisaLivro
End Sub
Private Sub cmd_sair_Click()
    Finaliza
    Unload Me
End Sub
Private Sub Form_Activate()
    If Empresa.LocalizarCodigo(g_empresa) = False Then
        Finaliza
        Unload Me
    End If
    If lChamaPesquisa = True Then
        PesquisaLivro
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_imprimir_Click
    End If
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    Call GravaAuditoria(1, Me.name, 1, "")
    
    lChamaPesquisa = True
    PreencheCboCombustivel
    l_local = 1
End Sub
Private Sub msk_data_f_GotFocus()
    msk_data_f.SelStart = 0
    msk_data_f.SelLength = 5
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_numero_livro.SetFocus
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
Private Sub txt_numero_livro_GotFocus()
    txt_numero_livro.SelStart = 0
    txt_numero_livro.SelLength = Len(txt_numero_livro.Text)
End Sub
Private Sub txt_numero_livro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_combustivel.SetFocus
    End If
    'Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_numero_livro_LostFocus()
    'txt_numero_livro.Text = Format(Val(txt_numero_livro.Text), "000")
    If chkNumeraTermos.Value = 1 Then
        txtPaginaFinal.Text = Format(CDate(msk_data_f.Text) - CDate(msk_data_i.Text) + 3, "000")
    Else
        txtPaginaInicial.Text = "001"
        txtPaginaFinal.Text = Format(CDate(msk_data_f.Text) - CDate(msk_data_i.Text) + 1, "000")
    End If
End Sub
Private Sub txtPaginaFinal_GotFocus()
    txtPaginaFinal.SelStart = 0
    txtPaginaFinal.SelLength = Len(txtPaginaFinal.Text)
End Sub
Private Sub txtPaginaFinal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtPaginaInicial_GotFocus()
    txtPaginaInicial.SelStart = 0
    txtPaginaInicial.SelLength = Len(txtPaginaInicial.Text)
End Sub
Private Sub txtPaginaInicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtPaginaFinal.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
