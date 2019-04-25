VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form lst_cheque_deposito_bancario 
   Caption         =   "Emissão do Sumário p/ Depósito Bancário"
   ClientHeight    =   2775
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_cheque_deposito_bancario.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_cheque_deposito_bancario.frx":030A
   ScaleHeight     =   2775
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_cheque_deposito_bancario.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_cheque_deposito_bancario.frx":19E2
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Imprime o sumário de cheques para depósito bancário."
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_cheque_deposito_bancario.frx":2FEC
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Visualiza o sumário de cheques para depósito bancário."
      Top             =   1800
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "lst_cheque_deposito_bancario.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_cheque_deposito_bancario.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_cheque_deposito_bancario.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_remessa 
         Height          =   315
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   11
         Top             =   1080
         Width           =   675
      End
      Begin VB.TextBox txt_bordero 
         Height          =   315
         Left            =   4860
         MaxLength       =   3
         TabIndex        =   13
         Top             =   1080
         Width           =   675
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
      Begin VB.Label Label4 
         Caption         =   "&Remessa"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "&Borderô"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   12
         Top             =   1080
         Width           =   975
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
         Index           =   0
         Left            =   3840
         TabIndex        =   7
         Top             =   660
         Width           =   975
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
      Left            =   180
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_cheque_deposito_bancario"
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
Dim l_data As Date
Dim lSubTotal As Currency
Dim lTotal As Currency
Dim lSubQtd As Currency
Dim lTotalQtd As Currency
Dim l_bordero As Integer
Dim tbl_movimento_cheque_avista As Table

Private MovCheque As New cMovimentoCheque
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_movimento_cheque_avista.Close
    Set MovCheque = Nothing
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lSubTotal = 0
    lTotal = 0
    lSubQtd = 0
    lTotalQtd = 0
End Sub
Private Sub RelatorioAVista()
    Dim i As Integer
    Dim x_total As Currency
    ZeraVariaveis
    For i = 1 To 11
        If TotalizaChequeAVista(i) > 0 Then
            l_bordero = l_bordero + 1
            Call ImpDetAVista(i)
        End If
    Next
    If lTotal > 0 Then
        ImpTotalAVista
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Cheque p/ Depósito Bancário|@|"
        frm_preview.Show 1
    End If
    cmd_sair.SetFocus
End Sub
Private Sub RelatorioPreDatado()
    Dim i As Integer
    Dim x_total As Currency
    ZeraVariaveis
    l_bordero = CLng(txt_bordero) - 1
    For i = 1 To 11
        If TotalizaChequePreDatado(i) > 0 Then
            l_bordero = l_bordero + 1
            Call ImpDetPreDatado
        End If
    Next
    If lTotal > 0 Then
        ImpTotalPreDatado
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Cheque p/ Depósito Bancário|@|"
        frm_preview.Show 1
    End If
    cmd_sair.SetFocus
End Sub
Private Sub ImpDetAVista(x_empresa As Integer)
    Dim x_linha As String
    Dim i As Integer
    If lPagina = 0 Then
        ImpCabAVista
    End If
    If lLinha >= 55 Then
        x_linha = "+------+------+------+------+----------------+----------------+----------------+"
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.NewPage"
        ImpCabAVista
    End If
    x_linha = "|      |      |      |      |                |                |                |"
    Mid(x_linha, 3, 4) = "0638"
    Mid(x_linha, 10, 4) = Format(Val(txt_remessa), "0000")
    Mid(x_linha, 17, 4) = Format(x_empresa, "0000")
    i = Len(Format(lSubQtd, "###0"))
    Mid(x_linha, 24 + 4 - i, 4) = Format(lSubQtd, "###0")
    i = Len(Format(lSubTotal, "###,###,##0.00"))
    Mid(x_linha, 31 + 14 - i, i) = Format(lSubTotal, "###,###,##0.00")
    i = Len(Format(lSubTotal, "###,###,##0.00"))
    Mid(x_linha, 48 + 14 - i, i) = Format(lSubTotal, "###,###,##0.00")
    i = Len(Format(0, "###,###,##0.00"))
    Mid(x_linha, 65 + 14 - i, i) = Format(0, "###,###,##0.00")
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpDetPreDatado()
    Dim x_linha As String
    Dim i As Integer
    If lPagina = 0 Then
        ImpCabPreDatado
    End If
    If lLinha >= 55 Then
        x_linha = "+------+------+------+------+----------------+----------------+----------------+"
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.NewPage"
        ImpCabPreDatado
    End If
    x_linha = "|      |      |      |      |                |                |                |"
    Mid(x_linha, 3, 4) = "4436"
    Mid(x_linha, 10, 4) = Format(Val(txt_remessa), "0000")
    Mid(x_linha, 17, 4) = Format(l_bordero, "0000")
    i = Len(Format(lSubQtd, "###0"))
    Mid(x_linha, 24 + 4 - i, 4) = Format(lSubQtd, "###0")
    i = Len(Format(lSubTotal, "###,###,##0.00"))
    Mid(x_linha, 31 + 14 - i, i) = Format(lSubTotal, "###,###,##0.00")
    i = Len(Format(lSubTotal, "###,###,##0.00"))
    Mid(x_linha, 48 + 14 - i, i) = Format(lSubTotal, "###,###,##0.00")
    i = Len(Format(0, "###,###,##0.00"))
    Mid(x_linha, 65 + 14 - i, i) = Format(0, "###,###,##0.00")
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotalAVista()
    Dim x_linha As String
    Dim i As Integer
    x_linha = "+------+------+------+------+----------------+----------------+----------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|        TOTAL GERAL |      |                |                |                |"
    i = Len(Format(lTotalQtd, "###0"))
    Mid(x_linha, 24 + 4 - i, 4) = Format(lTotalQtd, "###0")
    i = Len(Format(lTotal, "###,###,##0.00"))
    Mid(x_linha, 31 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
    i = Len(Format(lTotal, "###,###,##0.00"))
    Mid(x_linha, 48 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
    i = Len(Format(0, "###,###,##0.00"))
    Mid(x_linha, 65 + 14 - i, i) = Format(0, "###,###,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "+--------------------+------+----------------+----------------+----------------+"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & "  "
End Sub
Private Sub ImpTotalPreDatado()
    Dim x_linha As String
    Dim i As Integer
    x_linha = "+------+------+------+------+----------------+----------------+----------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|        TOTAL GERAL |      |                |                |                |"
    i = Len(Format(lTotalQtd, "###0"))
    Mid(x_linha, 24 + 4 - i, 4) = Format(lTotalQtd, "###0")
    i = Len(Format(lTotal, "###,###,##0.00"))
    Mid(x_linha, 31 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
    i = Len(Format(lTotal, "###,###,##0.00"))
    Mid(x_linha, 48 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
    i = Len(Format(0, "###,###,##0.00"))
    Mid(x_linha, 65 + 14 - i, i) = Format(0, "###,###,##0.00")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "+--------------------+------+----------------+----------------+----------------+"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & "  "
End Sub
Private Sub ImpCabAVista()
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
    x_string_40 = "GRUPO X          "
    g_string = ReadINI("GRUPO DE EMPRESAS", "Nome do Grupo", gArquivoIni)
    x_string_40 = g_string
    g_string = ""
    
    BioImprime "@Printer.Print " & "| " & x_string_40 & "                         Página, " & Format(lPagina, "000") & " |"
    BioImprime "@Printer.Print " & "| SUMÁRIO DE LOTES (CHEQUES A VISTA)                       Goiânia, " & msk_data & " |"
    BioImprime "@Printer.Print " & "| Referente a.: " & msk_data_i & " a " & msk_data_f & "                                        |"
    BioImprime "@Printer.Print " & "+------+------+------+------+----------------+----------------+----------------+"
    BioImprime "@Printer.Print " & "| AGEN | REM. | LOTE | DOCS.|   VALOR CAPA   | TOTAL DIGITADO |    DIFERENÇA   |"
    BioImprime "@Printer.Print " & "+------+------+------+------+----------------+----------------+----------------+"
End Sub
Private Sub ImpCabPreDatado()
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
    x_string_40 = "GRUPO X          "
    g_string = ReadINI("GRUPO DE EMPRESAS", "Nome do Grupo", gArquivoIni)
    x_string_40 = g_string
    g_string = ""
    BioImprime "@Printer.Print " & "| " & x_string_40 & "                         Página, " & Format(lPagina, "000") & " |"
    BioImprime "@Printer.Print " & "| SUMÁRIO DE BORDERÔ (CHEQUES PRE-DATADOS)                 Goiânia, " & msk_data & " |"
    BioImprime "@Printer.Print " & "| Referente a.: " & msk_data_i & " a " & msk_data_f & "                                        |"
    BioImprime "@Printer.Print " & "+------+------+------+------+----------------+----------------+----------------+"
    BioImprime "@Printer.Print " & "| POLO | REM. | BORD.| DOCS.|   VALOR CAPA   | TOTAL DIGITADO |    DIFERENÇA   |"
    BioImprime "@Printer.Print " & "+------+------+------+------+----------------+----------------+----------------+"
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        txt_remessa.SetFocus
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
    txt_remessa.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        txt_remessa.SetFocus
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
            RelatorioAVista
            RelatorioPreDatado
        End If
    End If
End Sub
Function TotalizaChequeAVista(x_empresa As Integer) As Currency
    Dim xQtd As Integer
    Dim xValor As Currency
    TotalizaChequeAVista = 0
    lSubTotal = 0
    lSubQtd = 0
    With tbl_movimento_cheque_avista
        If .RecordCount > 0 Then
            .Seek ">=", x_empresa, CDate(msk_data_i), " ", " ", 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> x_empresa Or ![Data de Emissao] > CDate(msk_data_f) Then
                        Exit Do
                    End If
                    TotalizaChequeAVista = TotalizaChequeAVista + !Valor
                    lSubTotal = lSubTotal + !Valor
                    lTotal = lTotal + !Valor
                    lSubQtd = lSubQtd + 1
                    lTotalQtd = lTotalQtd + 1
                    .MoveNext
                Loop
            End If
        End If
    End With
    
    xValor = MovCheque.TotalEmissaoPeriodo(x_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), "1", "9", "0", "V")
    xQtd = MovCheque.TotalQtdEmissaoPeriodo(x_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), "1", "9", "0", "V")
    TotalizaChequeAVista = TotalizaChequeAVista + xValor
    lSubTotal = lSubTotal + xValor
    lTotal = lTotal + xValor
    lSubQtd = lSubQtd + xQtd
    lTotalQtd = lTotalQtd + xQtd
End Function
Function TotalizaChequePreDatado(x_empresa As Integer) As Currency
    lSubTotal = MovCheque.TotalEmissaoPeriodo(x_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), "1", "9", "0", "P")
    lSubQtd = MovCheque.TotalQtdEmissaoPeriodo(x_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), "1", "9", "0", "P")
    TotalizaChequePreDatado = lSubTotal
    lTotal = lTotal + lSubTotal
    lTotalQtd = lTotalQtd + lSubQtd
End Function
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
    ElseIf Not Val(txt_remessa) > 0 Then
        MsgBox "Informe o número da remessa.", 64, "Atenção!"
        txt_remessa.SetFocus
    ElseIf Not Val(txt_bordero) > 0 Then
        MsgBox "Informe o número do borderô.", 64, "Atenção!"
        txt_bordero.SetFocus
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
            RelatorioAVista
            RelatorioPreDatado
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        txt_remessa.Text = "1"
        txt_bordero.Text = 1
        cmd_imprimir.SetFocus
        Screen.MousePointer = 1
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
    CentraForm Me
    Set tbl_movimento_cheque_avista = bd_sgp.OpenTable("Movimento_Cheque_Avista")
    tbl_movimento_cheque_avista.Index = "id_digitacao"
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
        txt_remessa.SetFocus
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
Private Sub txt_bordero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_bordero_LostFocus()
    txt_bordero = Format(txt_bordero, "000")
End Sub
Private Sub txt_remessa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_bordero.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
