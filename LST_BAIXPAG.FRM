VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_baixa_pagar 
   Caption         =   "Relação da Baixa dos Vencimentos"
   ClientHeight    =   4605
   ClientLeft      =   1545
   ClientTop       =   2970
   ClientWidth     =   6795
   Icon            =   "LST_BAIXPAG.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "LST_BAIXPAG.frx":030A
   ScaleHeight     =   4605
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1200
      Picture         =   "LST_BAIXPAG.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Visualiza baixa de contas à pagar."
      Top             =   3600
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "LST_BAIXPAG.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Imprime baixa de contas à pagar."
      Top             =   3600
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4800
      Picture         =   "LST_BAIXPAG.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3600
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CheckBox chkUnificaEmpresa 
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   2760
         Width           =   435
      End
      Begin VB.ComboBox cbo_tipo_despesa 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2340
         Width           =   3435
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "LST_BAIXPAG.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2700
         Picture         =   "LST_BAIXPAG.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2700
         Picture         =   "LST_BAIXPAG.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin VB.ComboBox cbo_fornecedor 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1920
         Width           =   4755
      End
      Begin VB.ComboBox cbo_conta 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1500
         Width           =   3435
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
         Caption         =   "&Unifica empresas"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo de Despesa"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   2340
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "Fornecedor"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Conta"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Data Inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Data Final"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Data de Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin Crystal.CrystalReport cr_relato 
      Left            =   120
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\VB5\Sgp\Data\LST_BAIXPAG.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   780
      Top             =   3780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_baixa_pagar"
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

Dim rstBaixaContasPagar As ADODB.Recordset
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
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
    x_linha = "| RELAÇÃO DAS BAIXAS DE VENCIMENTOS                               , __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(x_linha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(x_linha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| PERÍODO DE REFERÊNCIA...: __/__/____ A __/__/____                            |"
    Mid(x_linha, 29, 10) = msk_data_i.Text
    Mid(x_linha, 42, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| CONTA...................:                                                    |"
    Mid(x_linha, 29, 40) = cbo_conta.Text
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| FORNECEDOR..............:                                                    |"
    Mid(x_linha, 29, 40) = cbo_fornecedor.Text
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+------------------------------------------+----------+---------------+------------+------------+----------+------------------------+---+"
    BioImprime "@Printer.Print " & "| NOME DO FORNECEDOR                       | DATA  DO |  VALOR  PAGO  | NUMERO  DO | LOCAL   DE | DATA  DA | COMPLEMENTO            |EMP|"
    BioImprime "@Printer.Print " & "|                                          | PAGAMENTO|               |  DOCUMENTO |  COBRANCA  |  EMISSAO |                        |   |"
    BioImprime "@Printer.Print " & "+------------------------------------------+----------+---------------+------------+------------+----------+------------------------+---+"

End Sub
Private Sub ImpDet()
    Dim xLinha As String
    Dim i As Integer
    If lPagina = 0 Then
        ImpCab
    End If
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    If lLinha >= 60 Then
        xLinha = "+------------------------------------------+----------+---------------+------------+------------+----------+------------------------+---+"
        Mid(xLinha, 73, 20) = " Cerrado Tecnologia "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    '                  1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    xLinha = "|                                          |          |               |            |            |          |                        |   |"
    
    Mid(xLinha, 3, 40) = rstBaixaContasPagar!Nome_Fornecedor
    Mid(xLinha, 45, 10) = Format(rstBaixaContasPagar!Data_Pagamento, "dd/mm/yyyy")
    i = Len(Format(rstBaixaContasPagar!Valor_Pagamento, "######,##0.00"))
    Mid(xLinha, 57 + 13 - i, i) = Format(rstBaixaContasPagar!Valor_Pagamento, "######,##0.00")
    Mid(xLinha, 73, 10) = rstBaixaContasPagar!Numero_Documento
    Mid(xLinha, 86, 10) = rstBaixaContasPagar!NomeLocalCobranca
    Mid(xLinha, 98, 10) = Format(rstBaixaContasPagar!Data_Emissao, "dd/mm/yyyy")
    Mid(xLinha, 109, 24) = rstBaixaContasPagar!Complemento
    Mid(xLinha, 135, 2) = rstBaixaContasPagar!Empresa
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    BioImprime "@Printer.Print " & "+------------------------------------------+----------+---------------+------------+------------+----------+------------------------+---+"
    xLinha = "|                                       ***  TOTAL    |               |                                                                 |"
    i = Len(Format(lTotal, "######,##0.00"))
    Mid(xLinha, 57 + 13 - i, i) = Format(lTotal, "######,##0.00")
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    Mid(xLinha, 5, 20) = " Cerrado Tecnologia "
    xLinha = "+-----------------------------------------------------+---------------+-----------------------------------------------------------------+"
    Mid(xLinha, 73, 20) = " Cerrado Tecnologia "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub LoopBaixaContasPagar()
    Dim x_local_cobranca As String
    
    lSQL = ""
    lSQL = lSQL & "SELECT Baixa_Pagar.Empresa, Nome_Fornecedor, Data_Pagamento, Valor_Pagamento, Numero_Documento, Local_Cobranca, Data_Emissao, Numero_Documento, Complemento, Local_Cobrancas.Nome AS NomeLocalCobranca" 'new 10/02
    lSQL = lSQL & "  FROM Baixa_Pagar, Local_Cobrancas"
    lSQL = lSQL & " WHERE Data_Pagamento >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "   AND Data_Pagamento <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "   AND Local_Cobranca = Local_Cobrancas.Codigo"
    lSQL = lSQL & "   AND Baixa_Pagar.Empresa = Local_Cobrancas.Empresa"
    If chkUnificaEmpresa.Value = 0 Then
        lSQL = lSQL & "    AND Baixa_Pagar.Empresa = " & g_empresa
    End If
    If cbo_conta.ItemData(cbo_conta.ListIndex) > 0 Then
        lSQL = lSQL & " AND Codigo_Conta = " & cbo_conta.ItemData(cbo_conta.ListIndex)
    End If
    If cbo_fornecedor.ItemData(cbo_fornecedor.ListIndex) > 0 Then
        lSQL = lSQL & " And Codigo_Fornecedor = " & cbo_fornecedor.ItemData(cbo_fornecedor.ListIndex)
    End If
    If cbo_tipo_despesa.ItemData(cbo_tipo_despesa.ListIndex) > 0 Then
        lSQL = lSQL & " And [Tipo de Baixa] = " & Val(cbo_tipo_despesa.ItemData(cbo_tipo_despesa.ListIndex))
    End If
    lSQL = lSQL & " ORDER BY Data_Pagamento, Nome_Fornecedor, Registro"
    Set rstBaixaContasPagar = Conectar.RsConexao(lSQL)
        
    If rstBaixaContasPagar.RecordCount > 0 Then
        Do Until rstBaixaContasPagar.EOF
            lTotal = lTotal + rstBaixaContasPagar!Valor_Pagamento
            Call ImpDet
            rstBaixaContasPagar.MoveNext
        Loop
    End If
    rstBaixaContasPagar.Close
    Set rstBaixaContasPagar = Nothing
End Sub
Private Sub PreencheCboConta()
    Dim rstConta As New ADODB.Recordset
    
    cbo_conta.Clear
    cbo_conta.AddItem "Todas as Contas"
    cbo_conta.ItemData(cbo_conta.NewIndex) = 0
    Set rstConta = Conectar.RsConexao("SELECT Codigo, Nome FROM contas WHERE Empresa = " & g_empresa & "  ORDER BY NOME")
    Do Until rstConta.EOF
        cbo_conta.AddItem rstConta!Nome
        cbo_conta.ItemData(cbo_conta.NewIndex) = rstConta!Codigo
        rstConta.MoveNext
    Loop
    rstConta.Close
    Set rstConta = Nothing
End Sub
Private Sub PreencheCboTipoDespesa()
    cbo_tipo_despesa.Clear
    cbo_tipo_despesa.AddItem "Todas as Despesas"
    cbo_tipo_despesa.ItemData(cbo_tipo_despesa.NewIndex) = 0
    cbo_tipo_despesa.AddItem "Baixa Contas à Pagar"
    cbo_tipo_despesa.ItemData(cbo_tipo_despesa.NewIndex) = 1
    cbo_tipo_despesa.AddItem "Despesas de Caixa"
    cbo_tipo_despesa.ItemData(cbo_tipo_despesa.NewIndex) = 2
    cbo_tipo_despesa.AddItem "Falta de Caixa"
    cbo_tipo_despesa.ItemData(cbo_tipo_despesa.NewIndex) = 3
    cbo_tipo_despesa.AddItem "Vale de Funcionário"
    cbo_tipo_despesa.ItemData(cbo_tipo_despesa.NewIndex) = 4
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    Call LoopBaixaContasPagar
    If lTotal > 0 Then
        Call ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Baixa de Vencimentos|@|"
        frm_preview.Show 1
    Else
        MsgBox "Não existe baixa de vencimento no período informado!", vbInformation, "Atenção!"
    End If
    cmd_sair.SetFocus

'**** rotina do cristal report
'    Dim x_data_i As String
'    Dim x_data_f As String
'    x_data_i = "date(" & Format(msk_data_i, "yyyy,mm,dd") & ")"
'    x_data_f = "date(" & Format(msk_data_f, "yyyy,mm,dd") & ")"
'    cr_relato.SortFields(0) = "+{baixa_pagar.data_pagamento}"
'    cr_relato.SortFields(1) = "+{baixa_pagar.nome_fornecedor}"
'    cr_relato.SortFields(2) = "+{baixa_pagar.registro}"
'    cr_relato.Formulas(0) = "f_empresa = BeforeReadingRecords;""" & g_nome_empresa & """"
'    cr_relato.Formulas(1) = "f_data_hoje =  BeforeReadingRecords;""" & msk_data & """"
'    cr_relato.Formulas(2) = "f_data_inicial =  BeforeReadingRecords;""" & msk_data_i & """"
'    cr_relato.Formulas(3) = "f_data_final =  BeforeReadingRecords;""" & msk_data_f & """"
'    cr_relato.Formulas(4) = "f_conta =  BeforeReadingRecords;""" & cbo_conta.Text & """"
'    cr_relato.Formulas(5) = "f_fornecedor =  BeforeReadingRecords;""" & cbo_fornecedor.Text & """"
'    cr_relato.Formulas(6) = "f_tipo_despesa =  BeforeReadingRecords;""" & cbo_tipo_despesa.Text & """"
'    l_sql = "{baixa_pagar.empresa} = " & g_empresa
'    l_sql = l_sql & " And {baixa_pagar.data_pagamento} >= " & x_data_i
'    l_sql = l_sql & " And {baixa_pagar.data_pagamento} <= " & x_data_f
'    If cbo_conta.ItemData(cbo_conta.ListIndex) > 0 Then
'        l_sql = l_sql & " And {baixa_pagar.codigo_conta} = " & cbo_conta.ItemData(cbo_conta.ListIndex)
'    End If
'    If cbo_fornecedor.ItemData(cbo_fornecedor.ListIndex) > 0 Then
'        l_sql = l_sql & " And {baixa_pagar.codigo_fornecedor} = " & cbo_fornecedor.ItemData(cbo_fornecedor.ListIndex)
'    End If
'    If cbo_tipo_despesa.ItemData(cbo_tipo_despesa.ListIndex) > 0 Then
'        l_sql = l_sql & " And {baixa_pagar.Tipo de Baixa} = " & Val(cbo_tipo_despesa.ItemData(cbo_tipo_despesa.ListIndex))
'    End If
'    cr_relato.SelectionFormula = l_sql
'    cr_relato.Action = 1
End Sub
Private Sub cbo_conta_GotFocus()
    SendMessageLong cbo_conta.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_conta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_fornecedor.SetFocus
    End If
End Sub
Private Sub cbo_fornecedor_GotFocus()
    SendMessageLong cbo_fornecedor.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_fornecedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo_despesa.SetFocus
    End If
End Sub
Private Sub cbo_tipo_despesa_GotFocus()
    SendMessageLong cbo_tipo_despesa.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_tipo_despesa_KeyPress(KeyAscii As Integer)
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
        cbo_conta.SetFocus
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
    cbo_conta.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cbo_conta.SetFocus
    Else
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_imprimir_Click()
    If ValidaCampos Then
        If SelecionaImpressoraEpson(Me) Then
            cr_relato.Destination = 1
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
            cmd_sair.SetFocus
        End If
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
    ElseIf cbo_conta.ListIndex = -1 Then
        MsgBox "Selecione uma conta.", vbInformation, "Atenção!"
        cbo_conta.SetFocus
    ElseIf cbo_tipo_despesa.ListIndex = -1 Then
        MsgBox "Selecione um tipo de despesa.", vbInformation, "Atenção!"
        cbo_tipo_despesa.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lTotal = 0
End Sub
Private Sub cmd_sair_Click()
    Finaliza
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    If ValidaCampos Then
        If SelecionaImpressoraEpson(Me) Then
            cr_relato.Destination = 0
            Call GravaAuditoria(1, Me.name, 6, "")
            Relatorio
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        msk_data_i.Text = Format(Date - 30, "dd/mm/yyyy")
        msk_data_f.Text = Format(Date, "dd/mm/yyyy")
        cbo_conta.ListIndex = 0
        cbo_fornecedor.ListIndex = 0
        cbo_tipo_despesa.ListIndex = 0
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
    PreencheCboConta
    PreencheCboFornecedor
    PreencheCboTipoDespesa
End Sub
Private Sub PreencheCboFornecedor()
    Dim rstFornecedor As New ADODB.Recordset
    
    cbo_fornecedor.Clear
    cbo_fornecedor.AddItem "Todos os Fornecedores"
    cbo_fornecedor.ItemData(cbo_fornecedor.NewIndex) = 0
    Set rstFornecedor = Conectar.RsConexao("SELECT Codigo, Nome FROM Fornecedor ORDER BY NOME")
    Do Until rstFornecedor.EOF
        cbo_fornecedor.AddItem rstFornecedor!Nome
        cbo_fornecedor.ItemData(cbo_fornecedor.NewIndex) = rstFornecedor!Codigo
        rstFornecedor.MoveNext
    Loop
    rstFornecedor.Close
    Set rstFornecedor = Nothing
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
