VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_caixa 
   Caption         =   "Emissão do Caixa"
   ClientHeight    =   2355
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_caixa.frx":0000
   ScaleHeight     =   2355
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_caixa.frx":0046
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Visualiza o caixa."
      Top             =   1380
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_caixa.frx":1320
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprime o caixa."
      Top             =   1380
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_caixa.frx":25FA
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
      Width           =   6555
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "lst_caixa.frx":38D4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_caixa.frx":4BAE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_caixa.frx":5E88
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4860
         TabIndex        =   8
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
         Left            =   3960
         TabIndex        =   7
         Top             =   660
         Width           =   855
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
      Left            =   0
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_caixa"
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
Dim lSaldoAbertura As Currency
Dim lDebito As Currency
Dim lCredito As Currency
Dim lSaldoAtual As Currency
Dim tbl_empresa As Table
Dim tbl_movimento_caixa As Table
Dim tbl_saldo_caixa As Table
Private Sub Finaliza()
    tbl_empresa.Close
    tbl_movimento_caixa.Close
    tbl_saldo_caixa.Close
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lSaldoAbertura = 0
    lDebito = 0
    lCredito = 0
    lSaldoAtual = 0
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'Verifica movimento de caixa
    With tbl_movimento_caixa
        If .RecordCount > 0 Then
            .Seek ">", g_empresa, CDate(msk_data_i), 0
            If Not .NoMatch Then
                If !Empresa = g_empresa And !Data <= CDate(msk_data_f) Then
                    ImpDados
                    cmd_sair.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End With
    MsgBox "Não existe movimento de caixa no peíodo informado.", vbInformation, "Mensagem do Sistema"
End Sub
Private Sub ImpDados()
    'loop movimento de caixa
    With tbl_saldo_caixa
        If .RecordCount > 0 Then
            .Seek "<", g_empresa, CDate(msk_data_i)
            If Not .NoMatch Then
                lSaldoAbertura = !Saldo
                lSaldoAtual = !Saldo
            End If
        End If
    End With
    With tbl_movimento_caixa
        Do Until .EOF
            If !Empresa <> g_empresa Or !Data > CDate(msk_data_f) Then
                Exit Do
            End If
            ImpDet
            .MoveNext
        Loop
    End With
    If lDebito > 0 Or lCredito > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório do Livro Caixa|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub ImpDet()
    Dim x_linha As String
    Dim i As Integer
    With tbl_movimento_caixa
        If lPagina = 0 Then
            ImpCab
        End If
        If lLinha >= 57 Then
            x_linha = "+------------+------------+----------------------------------------------------+------------------+------------------+------------------+"
            Mid(x_linha, 31, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & x_linha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        If lPagina = 1 And lLinha = 0 Then
            x_linha = "|            |            |                                                    |                  |                  |                  |"
            Mid(x_linha, 29, 50) = "*** SALDO DE ABERTURA"
            i = Len(Format(lSaldoAbertura, "#,###,###,##0.00"))
            Mid(x_linha, 120 + 16 - i, i) = Format(lSaldoAbertura, "#,###,###,##0.00")
            BioImprime "@Printer.Print " & x_linha
            lLinha = lLinha + 1
        End If
        x_linha = "         1         2         3         4         5         6         7         8         9        10        11        12        13     13"
        x_linha = "12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567"
        x_linha = "|            |            |                                                    |                  |                  |                  |"
        Mid(x_linha, 3, 10) = Format(!Data, "dd/mm/yyyy")
        i = Len(Format(![Numero do Movimento], "##,###,##0"))
        Mid(x_linha, 16 + 10 - i, i) = Format(![Numero do Movimento], "##,###,##0")
        Mid(x_linha, 29, 50) = !Historico
        If ![Debito ou Credito] = "C" Then
            i = Len(Format(!valor, "#,###,###,##0.00"))
            Mid(x_linha, 82 + 16 - i, i) = Format(!valor, "#,###,###,##0.00")
            lCredito = lCredito + !valor
            lSaldoAtual = lSaldoAtual + !valor
        Else
            i = Len(Format(!valor, "#,###,###,##0.00"))
            Mid(x_linha, 101 + 16 - i, i) = Format(!valor, "#,###,###,##0.00")
            lDebito = lDebito + !valor
            lSaldoAtual = lSaldoAtual - !valor
        End If
        i = Len(Format(lSaldoAtual, "#,###,###,##0.00"))
        Mid(x_linha, 120 + 16 - i, i) = Format(lSaldoAtual, "#,###,###,##0.00")
    End With
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim x_linha As String
    Dim i As Integer
    x_linha = "+------------+------------+----------------------------------------------------+------------------+------------------+------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                                                                              |                                     |                  |"
    Mid(x_linha, 82, 50) = "*** SALDO DE ABERTURA"
    i = Len(Format(lSaldoAbertura, "#,###,###,##0.00"))
    Mid(x_linha, 120 + 16 - i, i) = Format(lSaldoAbertura, "#,###,###,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                                                                              |                                     |                  |"
    Mid(x_linha, 82, 50) = "*** TOTAL DE ENTRADA"
    i = Len(Format(lCredito, "#,###,###,##0.00"))
    Mid(x_linha, 120 + 16 - i, i) = Format(lCredito, "#,###,###,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                                                                              |                                     |                  |"
    Mid(x_linha, 82, 50) = "*** TOTAL DE SAIDA"
    i = Len(Format(lDebito, "#,###,###,##0.00"))
    Mid(x_linha, 120 + 16 - i, i) = Format(lDebito, "#,###,###,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                                                                              |                                     |                  |"
    Mid(x_linha, 82, 50) = "*** SALDO ATUAL"
    i = Len(Format(lSaldoAtual, "#,###,###,##0.00"))
    Mid(x_linha, 120 + 16 - i, i) = Format(lSaldoAtual, "#,###,###,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "+------------------------------------------------------------------------------+-------------------------------------+------------------+"
    Mid(x_linha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & "  "
End Sub
Private Sub ImpCab()
    Dim x_linha As String
    Dim x_cgc As String
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
    x_linha = "+------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                                                                  PÁGINA,     |"
    Mid(x_linha, 3, 40) = UCase(g_nome_empresa)
    Mid(x_linha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "         1         2         3         4         5         6         7         8         9        10        11        12        13     13"
    x_linha = "12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567"
    x_linha = "| C.G.C.:                            INSCRICAO ESTADUAL.:                      |"
    tbl_empresa.Seek "=", g_empresa
    If Not tbl_empresa.NoMatch Then
        x_cgc = Mid(tbl_empresa!CGC, 1, 2) & "." & Mid(tbl_empresa!CGC, 3, 3) & "." & Mid(tbl_empresa!CGC, 6, 3) & "/" & Mid(tbl_empresa!CGC, 9, 4) & "-" & Mid(tbl_empresa!CGC, 13, 2)
        Mid(x_linha, 11, 20) = x_cgc
        Mid(x_linha, 59, 20) = tbl_empresa![Inscricao Estadual]
    End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| LIVRO CAIXA                                              GOIÂNIA,            |"
    Mid(x_linha, 69, 10) = msk_data
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| REFERENTE A.:            a                                                   |"
    Mid(x_linha, 17, 10) = msk_data_i
    Mid(x_linha, 30, 10) = msk_data_f
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+------------+------------+----------------------------------------------------+------------------+------------------+------------------+"
    BioImprime "@Printer.Print " & "|  D A T A   |  NUMERO DO | HISTORICO                                          |   E N T R A D A  |     S A I D A    |     S A L D O    |"
    BioImprime "@Printer.Print " & "|            |  MOVIMENTO |                                                    |                  |                  |     A T U A L    |"
    BioImprime "@Printer.Print " & "+------------+------------+----------------------------------------------------+------------------+------------------+------------------+"
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cmd_imprimir.SetFocus
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
    cmd_imprimir.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cmd_imprimir.SetFocus
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
        MsgBox "Data final deve ser maior que a data inicial.", 64, "Atenção!"
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
            Relatorio
        End If
    End If
End Sub
Private Sub Form_Activate()
    If Not IsDate(msk_data) Then
        msk_data = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i = Format(g_data_def - 1, "dd/mm/yyyy")
        msk_data_f = Format(g_data_def - 1, "dd/mm/yyyy")
        msk_data_i.SetFocus
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_Load()
    CentraForm Me
    Set tbl_empresa = bd_sgp.OpenTable("Empresas")
    Set tbl_movimento_caixa = bd_sgp_m.OpenTable("Movimento_Caixa")
    Set tbl_saldo_caixa = bd_sgp.OpenTable("Saldo_Caixa")
    tbl_empresa.Index = "id_codigo"
    tbl_movimento_caixa.Index = "id_data"
    tbl_saldo_caixa.Index = "id_data"
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
        cmd_imprimir.SetFocus
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
