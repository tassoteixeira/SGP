VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form emissao_conta_pagar_conferencia 
   Caption         =   "Emissão de Contas à Pagar (Conferência)"
   ClientHeight    =   2355
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   6675
   Icon            =   "lst_conta_pagar_conferencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2355
   ScaleWidth      =   6675
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_conta_pagar_conferencia.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Visualiza contas à pagar para conferência."
      Top             =   1380
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2940
      Picture         =   "lst_conta_pagar_conferencia.frx":1A24
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprime contas à pagar para conferência."
      Top             =   1380
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4740
      Picture         =   "lst_conta_pagar_conferencia.frx":302E
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1380
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6435
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5820
         Picture         =   "lst_conta_pagar_conferencia.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_conta_pagar_conferencia.frx":599A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_conta_pagar_conferencia.frx":6C74
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4680
         TabIndex        =   8
         Top             =   660
         Width           =   1095
         _ExtentX        =   1931
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
         Width           =   1095
         _ExtentX        =   1931
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
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   3780
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_conta_pagar_conferencia"
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
Dim l_total(1 To 12) As Currency
Dim tbl_conta As Table
Dim tbl_empresa As Table
Dim tbl_fornecedor As Table
Dim tbl_movimento_contas_pagar As Table
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_conta.Close
    tbl_empresa.Close
    tbl_fornecedor.Close
    tbl_movimento_contas_pagar.Close
End Sub
Private Sub ZeraVariaveis()
Dim i As Integer
    lLinha = 0
    lPagina = 0
    For i = 1 To 12
        l_total(i) = 0
    Next
End Sub
Private Sub Relatorio()
    Dim i As Integer
    ZeraVariaveis
    Call LoopEmpresa
    If l_total(12) > 0 Then
        Call ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório do Contas à Pagar (Conferência)|@|"
        frm_preview.Show 1
    Else
        MsgBox "Não existe movimento no período informado!", vbInformation, "Mensagem do Sistema"
    End If
    cmd_sair.SetFocus
End Sub
Private Sub ImpMovimentoContasPagar(i As Integer)
    Dim x_nome_conta As String
    With tbl_movimento_contas_pagar
        .Seek ">=", i, CDate(msk_data_i), 0
        If Not .NoMatch Then
            Do Until .EOF
                If !Empresa <> i Or ![Data da Digitacao] > CDate(msk_data_f) Then
                    Exit Do
                End If
                tbl_conta.Seek "=", !codigo_conta
                If Not tbl_conta.NoMatch Then
                    x_nome_conta = tbl_conta!Nome
                Else
                    x_nome_conta = "** Inexistente **"
                End If
                l_total(i) = l_total(i) + !Valor
                l_total(12) = l_total(12) + !Valor
                Call ImpDet(i, !Nome_Fornecedor, !Data_Emissao, !Data_Vencimento, !Valor, !Numero_Documento, !Complemento, x_nome_conta, ![Data da Digitacao])
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub LoopEmpresa()
    Dim i As Integer
    With tbl_empresa
        If .RecordCount > 0 And tbl_movimento_contas_pagar.RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                If !Codigo = 5 Then
                    .MoveNext
                End If
                If !Codigo > 12 Then
                    Exit Do
                End If
                i = !Codigo
                ImpMovimentoContasPagar i
                If l_total(i) > 0 Then
                    ImpSubTotal i
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub ImpSubTotal(x_empresa As Integer)
    Dim x_linha As String
    Dim i As Integer
    x_linha = "|  |                         |          |          |             |          |                              |          |                 |"
    i = Len(Format(x_empresa, "#0"))
    Mid(x_linha, 2 + 2 - i, i) = Format(x_empresa, "#0")
    Mid(x_linha, 5, 25) = "** Total da Empresa"
    i = Len(Format(l_total(x_empresa), "##,###,##0.00"))
    Mid(x_linha, 53 + 13 - i, i) = Format(l_total(x_empresa), "##,###,##0.00")
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "+--+-------------------------+----------+----------+-------------+----------+------------------------------+----------+-----------------+"
    BioImprime "@Printer.Print " & x_linha
End Sub
Private Sub ImpTotal()
    Dim x_linha As String
    Dim i As Integer
    x_linha = "|  |                                               |             |                                                                      |"
    Mid(x_linha, 5, 25) = "*** Total Geral"
    i = Len(Format(l_total(12), "##,###,##0.00"))
    Mid(x_linha, 53 + 13 - i, i) = Format(l_total(12), "##,###,##0.00")
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "+--+-----------------------------------------------+-------------+----------------------------------------------------------------------+"
    Mid(x_linha, 8, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpDet(x_empresa As Integer, x_fornecedor As String, x_data_emissao As Date, x_data_vencimento As Date, x_valor As Currency, x_numero_documento As String, x_historico As String, x_conta As String, x_data_digitacao As Date)
    Dim x_linha As String
    Dim i As Integer
    '                   1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '          12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    If lPagina = 0 Then
        ImpCab
    End If
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    If lLinha >= 64 Then
        x_linha = "+--+-------------------------+----------+----------+-------------+----------+------------------------------+----------+-----------------+"
        Mid(x_linha, 5, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    x_linha = "|  |                         |          |          |             |          |                              |          |                 |"
    i = Len(Format(x_empresa, "#0"))
    Mid(x_linha, 2 + 2 - i, i) = Format(x_empresa, "#0")
    Mid(x_linha, 5, 25) = x_fornecedor
    Mid(x_linha, 31, 10) = x_data_emissao
    Mid(x_linha, 42, 10) = x_data_vencimento
    i = Len(Format(x_valor, "##,###,##0.00"))
    Mid(x_linha, 53 + 13 - i, i) = Format(x_valor, "##,###,##0.00")
    Mid(x_linha, 67, 10) = x_numero_documento
    Mid(x_linha, 78, 30) = x_historico
    Mid(x_linha, 109, 10) = x_data_digitacao
    Mid(x_linha, 120, 17) = x_conta
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
    x_linha = "| GRUPO X                                                          Página, " & Format(lPagina, "000") & " |"
    g_string = ReadINI("GRUPO DE EMPRESAS", "Nome do Grupo", gArquivoIni)
    Mid(x_linha, 3, 40) = g_string
    g_string = ""
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| CONTAS À PAGAR (CONFERÊNCIA)                                    , __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(x_linha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(x_linha, 69, 10) = msk_data
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| PERÍODO DA DIGITAÇÃO....: __/__/____ A __/__/____                            |"
    Mid(x_linha, 29, 10) = msk_data_i
    Mid(x_linha, 42, 10) = msk_data_f
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+--+-------------------------+----------+----------+-------------+----------+------------------------------+----------+-----------------+"
    BioImprime "@Printer.Print " & "|EM|FORNECEDOR               | DATA  DA | DATA  DO |  VALOR  DO  | NÚMERO DO| HISTÓRICO                    | DATA  DA | TIPO DA CONTA   |"
    BioImprime "@Printer.Print " & "|PR|                         |  EMISSÃO |VENCIMENTO|  VENCIMENTO | DOCUMENTO|                              | DIGITAÇÃO|                 |"
    BioImprime "@Printer.Print " & "+--+-------------------------+----------+----------+-------------+----------+------------------------------+----------+-----------------+"
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
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cmd_visualizar.SetFocus
    Else
        msk_data_i = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = " "
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
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data de emissão.", 64, "Atenção!"
        msk_data.SetFocus
    ElseIf Not IsDate(msk_data_i) Then
        MsgBox "Informe a data inicial.", 64, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f) Then
        MsgBox "Informe a data final.", 64, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f) < CDate(msk_data_i) Then
        MsgBox "Data final deve ser maior ou igual a " & CDate(msk_data_i) & ".", 64, "Atenção!"
        msk_data_f.SetFocus
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
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 6, "")
            Relatorio
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        msk_data_i.Text = Format(CDate(g_data_def), "dd/mm/yyyy")
        msk_data_f.Text = Format(CDate(g_data_def), "dd/mm/yyyy")
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
    Set tbl_conta = bd_sgp.OpenTable("Contas")
    Set tbl_empresa = bd_sgp.OpenTable("Empresas")
    Set tbl_fornecedor = bd_sgp.OpenTable("Fornecedor")
    Set tbl_movimento_contas_pagar = bd_sgp.OpenTable("Contas_Pagar")
    tbl_conta.Index = "id_codigo"
    tbl_empresa.Index = "id_codigo"
    tbl_fornecedor.Index = "id_codigo"
    tbl_movimento_contas_pagar.Index = "id_data_digitacao"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
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
