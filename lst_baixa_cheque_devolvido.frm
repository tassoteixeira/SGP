VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_baixa_cheque_devolvido 
   Caption         =   "Emissão da Baixa de Cheque Devolvido"
   ClientHeight    =   3555
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_baixa_cheque_devolvido.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_baixa_cheque_devolvido.frx":030A
   ScaleHeight     =   3555
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_baixa_cheque_devolvido.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Visualiza baixa de cheque devolvido."
      Top             =   2580
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_baixa_cheque_devolvido.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprime baixa de cheque devolvido."
      Top             =   2580
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_baixa_cheque_devolvido.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2580
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.ComboBox cboSituacao 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1860
         Width           =   4755
      End
      Begin VB.TextBox txt_recebido_por 
         Height          =   300
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   11
         Top             =   1080
         Width           =   4215
      End
      Begin VB.CheckBox chk_detalhado 
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   1500
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_baixa_cheque_devolvido.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_baixa_cheque_devolvido.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "lst_baixa_cheque_devolvido.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
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
      Begin VB.Label Label6 
         Caption         =   "Sit&uação"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   1860
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "&Recebido por"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "De&talhado"
         Height          =   315
         Left            =   120
         TabIndex        =   12
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
      Left            =   0
      Top             =   2820
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_baixa_cheque_devolvido"
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
Dim lSubTotal As Currency
Dim lTotal As Currency
Dim lSubQtd As Currency
Dim lTotalQtd As Currency
Dim lSubDias As Currency
Dim lTotalDias As Currency
Dim lSQL As String
Private rsBaixaMovCheque As New adodb.Recordset
Private rsTabela As New adodb.Recordset
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set rsBaixaMovCheque = Nothing
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lSubTotal = 0
    lTotal = 0
    lSubQtd = 0
    lTotalQtd = 0
    lSubDias = 0
    lTotalDias = 0
End Sub
Private Sub PreencheCboSituacao()
    cboSituacao.Clear
    cboSituacao.AddItem "Todas as Situações"
    cboSituacao.ItemData(cboSituacao.NewIndex) = 0
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Nome, Codigo"
    lSQL = lSQL & "  FROM Situacao_Cheque_Devolvido"
    lSQL = lSQL & " ORDER BY Nome"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            cboSituacao.AddItem rsTabela("Nome").Value
            cboSituacao.ItemData(cboSituacao.NewIndex) = rsTabela("Codigo").Value
            rsTabela.MoveNext
        Loop
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    
    'Prepara SQL
    lSQL = "SELECT [Data de Digitacao], [Data de Emissao], [Numero da Conta], [Numero do Cheque], Periodo,"
    lSQL = lSQL & " [Tipo do Movimento], Valor, [Data do Vencimento], Emitente,"
    lSQL = lSQL & " [Data da Devolucao], [Motivo da Devolucao], [Ordem da Digitacao],"
    lSQL = lSQL & " Bancos.Nome as NomeBanco, Situacao_Cheque_Devolvido.Nome as NomeSituacao,"
    lSQL = lSQL & " [Recebido Por], [Data do Pagamento], [Valor Pago Dinheiro],"
    lSQL = lSQL & " [Valor Pago Cheque a Vista], [Valor Pago Cheque a Prazo]"
    lSQL = lSQL & " FROM Baixa_Cheque_Devolvido, Bancos, Situacao_Cheque_Devolvido"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " AND [Data do Pagamento] >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & " AND [Data do Pagamento] <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & " AND Bancos.Codigo = [Codigo do Banco]"
    lSQL = lSQL & " AND Situacao_Cheque_Devolvido.Codigo = Situacao"
    If cboSituacao.ItemData(cboSituacao.ListIndex) > 0 Then
        lSQL = lSQL & " AND Situacao = " & cboSituacao.ItemData(cboSituacao.ListIndex)
    End If
    If Trim(txt_recebido_por.Text) <> "" Then
        lSQL = lSQL & " AND [Recebido Por] = " & Chr(39) & Trim(txt_recebido_por.Text) & Chr(39)
    End If
    lSQL = lSQL & " ORDER BY Bancos.Nome, [Data do Pagamento], [Data da Devolucao], [Ordem da Digitacao]"
    'Abre RecordSet
    Set rsBaixaMovCheque = New adodb.Recordset
    Set rsBaixaMovCheque = Conectar.RsConexao(lSQL)
    
    
    'Verifica movimento
    If rsBaixaMovCheque.RecordCount > 0 Then
        ImpDados
    End If
    If rsBaixaMovCheque.State = 1 Then
        rsBaixaMovCheque.Close
    End If
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    Dim xLinha As String
    'loop movimento de cheques
    Do Until rsBaixaMovCheque.EOF
        If lPagina = 0 Then
            ImpCab
            lData = rsBaixaMovCheque("Data do Vencimento").Value
        End If
        If lLinha >= 57 Then
            xLinha = "+------------+-------+-----------+----------+-----+---------------+------------+------------------------------------------+-------------+"
            Mid(xLinha, 84, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & xLinha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        'lData <> rsBaixaMovCheque("Data do Vencimento").Value Then
        '    ImpSubTotal
        '    lData = rsBaixaMovCheque("Data do Vencimento").Value
        'End If
        ImpDet
        lSubTotal = lSubTotal + rsBaixaMovCheque("Valor Pago Dinheiro").Value + rsBaixaMovCheque("Valor Pago Cheque a Vista").Value + rsBaixaMovCheque("Valor Pago Cheque a Prazo").Value
        lTotal = lTotal + rsBaixaMovCheque("Valor Pago Dinheiro").Value + rsBaixaMovCheque("Valor Pago Cheque a Vista").Value + rsBaixaMovCheque("Valor Pago Cheque a Prazo").Value
        lSubQtd = lSubQtd + 1
        lTotalQtd = lTotalQtd + 1
        lSubDias = lSubDias + DateDiff("d", rsBaixaMovCheque("Data de Emissao").Value, rsBaixaMovCheque("Data do Vencimento").Value)
        lTotalDias = lTotalDias + DateDiff("d", rsBaixaMovCheque("Data de Emissao").Value, rsBaixaMovCheque("Data do Vencimento").Value)
        rsBaixaMovCheque.MoveNext
    Loop
    'ImpSubTotal
    If lTotal > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Baixa de Cheque Devolvido|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub ImpDet()
    Dim xLinha As String
    Dim i As Integer
    Dim xValorPago As Currency
    
    xLinha = "|            |       |           |          |     |               |            |                                          |             |"
    Mid(xLinha, 3, 10) = Format(rsBaixaMovCheque("Data de Emissao").Value, "dd/mm/yyyy")
    Mid(xLinha, 18, 1) = rsBaixaMovCheque("Periodo").Value
    i = Len(Format(rsBaixaMovCheque("Numero da Conta").Value, "##########"))
    Mid(xLinha, 24 + 10 - i, i) = Format(rsBaixaMovCheque("Numero da Conta").Value, "##########")
    i = Len(Format(rsBaixaMovCheque("Numero do Cheque").Value, "######"))
    Mid(xLinha, 38 + 6 - i, i) = Format(rsBaixaMovCheque("Numero do Cheque").Value, "######")
    i = Len(Format(DateDiff("d", rsBaixaMovCheque("Data de Emissao").Value, rsBaixaMovCheque("Data do Vencimento").Value), "##"))
    Mid(xLinha, 48 + 2 - i, i) = Format(DateDiff("d", rsBaixaMovCheque("Data de Emissao").Value, rsBaixaMovCheque("Data do Vencimento").Value), "##")
    xValorPago = rsBaixaMovCheque("Valor Pago Dinheiro").Value + rsBaixaMovCheque("Valor Pago Cheque a Vista").Value + rsBaixaMovCheque("Valor Pago Cheque a Prazo").Value
    i = Len(Format(xValorPago, "###,###,##0.00"))
    Mid(xLinha, 52 + 14 - i, i) = Format(xValorPago, "###,###,##0.00")
    Mid(xLinha, 69, 10) = Format(rsBaixaMovCheque("Data do Vencimento").Value, "dd/mm/yyyy")
    Mid(xLinha, 82, 40) = rsBaixaMovCheque("Emitente").Value
    Mid(xLinha, 125, 10) = Format(rsBaixaMovCheque("Data do Pagamento").Value, "dd/mm/yyyy")
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
    If chk_detalhado.Value = 1 Then
        xLinha = "|                                |          |     |               |            |                                          |             |"
        i = Len(Format(rsBaixaMovCheque("Valor").Value, "###,###,##0.00"))
        Mid(xLinha, 52 + 14 - i, i) = Format(rsBaixaMovCheque("Valor").Value, "###,###,##0.00")
        Mid(xLinha, 82, 40) = rsBaixaMovCheque("Recebido Por").Value
        Mid(xLinha, 125, 10) = Format(rsBaixaMovCheque("Data da Devolucao").Value, "dd/mm/yyyy")
        BioImprime "@Printer.Print " & xLinha
        xLinha = "|                                |          |                     |            |                                          |             |"
        Mid(xLinha, 3, 40) = rsBaixaMovCheque("NomeBanco").Value
        Mid(xLinha, 36, 30) = rsBaixaMovCheque("NomeSituacao").Value
        Mid(xLinha, 73, 1) = rsBaixaMovCheque("Tipo do Movimento").Value
        Mid(xLinha, 82, 20) = rsBaixaMovCheque("Motivo da Devolucao").Value
        Mid(xLinha, 125, 10) = Format(rsBaixaMovCheque("Data de Digitacao").Value, "dd/mm/yyyy")
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 2
        BioImprime "@Printer.Print " & "+------------+-------+-----------+----------+-----+---------------+------------+------------------------------------------+-------------+"
        lLinha = lLinha + 1
    End If
End Sub
Private Sub ImpSubTotal()
    Dim xLinha As String
    Dim i As Integer
    If lSubTotal > 0 Then
        xLinha = "|            |       |           |          |     |               |            |                                          |             |"
        Mid(xLinha, 35, 10) = "*** TOTAL "
        i = Len(Format(lSubDias / lSubQtd, "#0.00"))
        Mid(xLinha, 46 + 5 - i, i) = Format(lSubDias / lSubQtd, "#0.00")
        i = Len(Format(lSubTotal, "###,###,##0.00"))
        Mid(xLinha, 52 + 14 - i, i) = Format(lSubTotal, "###,###,##0.00")
        Mid(xLinha, 82, 17) = "NÚMERO DE CHEQUES"
        i = Len(Format(lSubQtd, "####"))
        Mid(xLinha, 104 + 4 - i, i) = Format(lSubQtd, "####")
        Mid(xLinha, 108, 4) = " EM "
        Mid(xLinha, 112, 10) = Format(lData, "dd/mm/yyyy")
        BioImprime "@@Printer.FontName = Sans Serif 17cpi"
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
        lSubTotal = 0
        lSubQtd = 0
        lSubDias = 0
    End If
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    If chk_detalhado.Value = False Then
        BioImprime "@Printer.Print " & "+------------+-------+-----------+----------+-----+---------------+------------+------------------------------------------+-------------+"
    End If
    xLinha = "|                                           |     |               |            |                                          |             |"
    Mid(xLinha, 35, 10) = "*** TOTAL "
    i = Len(Format(lTotalDias / lTotalQtd, "#0.00"))
    Mid(xLinha, 46 + 5 - i, i) = Format(lTotalDias / lTotalQtd, "#0.00")
    i = Len(Format(lTotal, "###,###,##0.00"))
    Mid(xLinha, 52 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
    Mid(xLinha, 82, 17) = "NÚMERO DE CHEQUES"
    i = Len(Format(lTotalQtd, "####"))
    Mid(xLinha, 104 + 4 - i, i) = Format(lTotalQtd, "####")
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+-------------------------------------------+-----+---------------+------------+------------------------------------------+-------------+"
    Mid(xLinha, 84, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & "  "
End Sub
Private Sub ImpCab()
    Dim i As Integer
    Dim xLinha As String
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
    xLinha = "|                                                                  Página, ___ |"
    Mid(xLinha, 3, 40) = g_nome_empresa
    Mid(xLinha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| RELAÇÃO DA BAIXA DE CHEQUES DEVOLVIDOS                    CIDADE, __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| Referente a.: __/__/____ a __/__/____                                        |"
    Mid(xLinha, 17, 10) = msk_data_i.Text
    Mid(xLinha, 30, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| Situacao....: x                                                              |"
    Mid(xLinha, 17, 30) = cboSituacao.Text
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+------------+-------+-----------+----------+-----+---------------+------------+------------------------------------------+-------------+"
    BioImprime "@Printer.Print " & "|DATA EMISSAO|PERIODO|N. DA CONTA| N.CHEQUE |PRAZO|  VALOR  PAGO  | VENCIMENTO | NOME DO EMITENTE                         |DT. PAGAMENTO|"
    If chk_detalhado.Value = 1 Then
        BioImprime "@Printer.Print " & "|                                |          |     |VALOR DO CHEQUE|            | RECEBIDO POR                             |DT. DEVOLUCAO|"
        BioImprime "@Printer.Print " & "| BANCO                          | SITUACAO                       | TIPO  MOV. | MOTIVO DA DEVOLUCAO                      |DT. DIGITACAO|"
    End If
    BioImprime "@Printer.Print " & "+------------+-------+-----------+----------+-----+---------------+------------+------------------------------------------+-------------+"
End Sub
Private Sub cboSituacao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cmd_visualizar.SetFocus
    Else
        msk_data.Text = RetiraGString(1)
        msk_data_i.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_data_f_Click()
    g_string = msk_data_f
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
    Else
        msk_data_f.Text = RetiraGString(1)
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
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
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
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        msk_data_i.SetFocus
        cboSituacao.ListIndex = 0
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
    PreencheCboSituacao
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
        txt_recebido_por.SetFocus
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
Private Sub txt_recebido_por_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub txt_recebido_por_LostFocus()
    txt_recebido_por.SelStart = 0
    txt_recebido_por.SelLength = Len(txt_recebido_por.Text)
End Sub
