VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_cheque 
   Caption         =   "Emissão de Cheque Pré-Datado"
   ClientHeight    =   4260
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_cheque.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_cheque.frx":030A
   ScaleHeight     =   4260
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_cheque.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Visualiza cheques pré-datado."
      Top             =   3300
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_cheque.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Imprime cheques pré-datado."
      Top             =   3300
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_cheque.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3300
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   3075
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CheckBox chkChequePredatado 
         Caption         =   "Cheque Pré-Datado"
         Height          =   315
         Left            =   4260
         TabIndex        =   26
         Top             =   2640
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkChequeVista 
         Caption         =   "Cheque a Vista"
         Height          =   315
         Left            =   1680
         TabIndex        =   25
         Top             =   2640
         Width           =   2175
      End
      Begin VB.TextBox txt_taxa_juros 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   20
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_cheque.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_cheque.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "lst_cheque.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.ComboBox cbo_tipo_movimento 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1860
         Width           =   2175
      End
      Begin VB.ComboBox cbo_periodo_f 
         Height          =   315
         Left            =   5940
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1440
         Width           =   495
      End
      Begin VB.ComboBox cbo_periodo_i 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1440
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
      Begin VB.OptionButton optEmissao 
         Caption         =   "Emissão"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   1080
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optVencimento 
         Caption         =   "Vencimento"
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo de Cheque"
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   2640
         Width           =   1515
      End
      Begin VB.Label Label8 
         Caption         =   "Taxa de &Juros"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   1515
      End
      Begin VB.Label Label7 
         Caption         =   "I&mprimir por"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "Período &final"
         Height          =   315
         Left            =   3840
         TabIndex        =   15
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "&Tipo de Movimento"
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   1860
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Período inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   1440
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
      Top             =   3540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_cheque"
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
Dim lSubDias As Currency
Dim lTotalDias As Currency
Dim lTotalJuros As Currency
Dim lSQL As String

Private rsCheque As New adodb.Recordset
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
End Sub
Private Sub PreencheCboPeriodo()
    cbo_periodo_i.Clear
    cbo_periodo_f.Clear
    cbo_periodo_i.AddItem 1
    cbo_periodo_f.AddItem 1
    cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = 1
    cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = 1
    cbo_periodo_i.AddItem 2
    cbo_periodo_f.AddItem 2
    cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = 2
    cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = 2
    cbo_periodo_i.AddItem 3
    cbo_periodo_f.AddItem 3
    cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = 3
    cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = 3
    cbo_periodo_i.AddItem 4
    cbo_periodo_f.AddItem 4
    cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = 4
    cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = 4
End Sub
Private Sub PreencheCboTipoMovimento()
    cbo_tipo_movimento.Clear
    cbo_tipo_movimento.AddItem "0 Todos os Caixas"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 0
    cbo_tipo_movimento.AddItem "1 Caixa de combustíveis"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 1
    cbo_tipo_movimento.AddItem "2 Caixa de óleo/diversos"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 2
    cbo_tipo_movimento.AddItem "3 - Cheque Inclusão"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 3
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
    lTotalJuros = 0
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'Verifica movimento
    'Verifica Movimento_Cheque
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT [Data do Vencimento], [Data de Emissao], Valor, Periodo, [Numero da Conta], [Numero do Cheque], Emitente, [Tipo do Movimento], [Dados do Abastecimento]"
    lSQL = lSQL & "  FROM Movimento_Cheque"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    If optEmissao.Value = True Then
        lSQL = lSQL & "   AND [Data de Emissao] >= " & preparaData(msk_data_i.Text)
        lSQL = lSQL & "   AND [Data de Emissao] <= " & preparaData(msk_data_f.Text)
    Else
        lSQL = lSQL & "   AND [Data do Vencimento] >= " & preparaData(msk_data_i.Text)
        lSQL = lSQL & "   AND [Data do Vencimento] <= " & preparaData(msk_data_f.Text)
    End If
    lSQL = lSQL & "   AND Periodo >= " & preparaTexto(cbo_periodo_i.Text)
    lSQL = lSQL & "   AND Periodo <= " & preparaTexto(cbo_periodo_f.Text)
    If Val(cbo_tipo_movimento.Text) > 0 Then
        lSQL = lSQL & "   AND [Tipo do Movimento] = " & preparaTexto(Val(cbo_tipo_movimento.Text))
    End If
    If chkChequeVista.Value = 0 Then
        lSQL = lSQL & "   AND [Data de Emissao] <> [Data do Vencimento]"
    Else
        If chkChequePredatado.Value = 0 Then
            lSQL = lSQL & "   AND [Data de Emissao] = [Data do Vencimento]"
        End If
    End If
    If optEmissao.Value = True Then
        lSQL = lSQL & " ORDER BY [Data de Emissao], Periodo, [Tipo do Movimento], [Ordem da Digitacao], [Numero da Conta], [Numero do Cheque]"
    Else
        lSQL = lSQL & " ORDER BY [Data do Vencimento], [Data de Emissao], Periodo, [Tipo do Movimento], [Ordem da Digitacao], [Numero da Conta], [Numero do Cheque]"
    End If
    'Abre RecordSet
    Set rsCheque = New adodb.Recordset
    Set rsCheque = Conectar.RsConexao(lSQL)
    If rsCheque.RecordCount > 0 Then
        ImpDados
    End If
    If rsCheque.State = 1 Then
        rsCheque.Close
    End If
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    Dim x_linha As String
    'loop movimento de cheques
    Do Until rsCheque.EOF
        If lPagina = 0 Then
            ImpCab
            l_data = rsCheque("Data do Vencimento").Value
        End If
        If lLinha >= 57 Then
            x_linha = "+------------+-------+-----------+----------+-----+---------------+------------+------------------------------------------+-------------+"
            Mid(x_linha, 84, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & x_linha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        If optVencimento.Value = True And l_data <> rsCheque("Data do Vencimento").Value Then
            ImpSubTotal
            l_data = rsCheque("Data do Vencimento").Value
        End If
        ImpDet
        lSubTotal = lSubTotal + rsCheque("valor").Value
        lTotal = lTotal + rsCheque("valor").Value
        lSubQtd = lSubQtd + 1
        lTotalQtd = lTotalQtd + 1
        lSubDias = lSubDias + DateDiff("d", rsCheque("Data de Emissao").Value, rsCheque("Data do Vencimento").Value)
        lTotalDias = lTotalDias + DateDiff("d", rsCheque("Data de Emissao").Value, rsCheque("Data do Vencimento").Value)
        rsCheque.MoveNext
    Loop
    If optVencimento.Value = True Then
        ImpSubTotal
    End If
    If lTotal > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Cheque Pré-Datado|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub ImpDet()
    Dim x_linha As String
    Dim i As Integer
    Dim xValor As Currency
    Dim xDias As Integer
    x_linha = "|            |       |           |          |     |               |            |                                          |             |"
    Mid(x_linha, 3, 10) = Format(rsCheque("Data de Emissao").Value, "dd/mm/yyyy")
    Mid(x_linha, 18, 1) = rsCheque("Periodo").Value
    i = Len(Format(rsCheque("Numero da Conta").Value, "##########"))
    Mid(x_linha, 24 + 10 - i, i) = Format(rsCheque("Numero da Conta").Value, "##########")
    i = Len(Format(rsCheque("Numero do Cheque").Value, "######"))
    Mid(x_linha, 38 + 6 - i, i) = Format(rsCheque("Numero do Cheque").Value, "######")
    If txt_taxa_juros.Text <> "" Then
        xDias = DateDiff("d", CDate(msk_data.Text), rsCheque("Data do Vencimento").Value)
    Else
        xDias = DateDiff("d", rsCheque("Data de Emissao").Value, rsCheque("Data do Vencimento").Value)
    End If
    i = Len(Format(xDias, "##"))
    Mid(x_linha, 48 + 2 - i, i) = xDias
    i = Len(Format(rsCheque("Valor").Value, "###,###,##0.00"))
    Mid(x_linha, 52 + 14 - i, i) = Format(rsCheque("Valor").Value, "###,###,##0.00")
    Mid(x_linha, 69, 10) = Format(rsCheque("Data do Vencimento").Value, "dd/mm/yyyy")
    Mid(x_linha, 82, 40) = rsCheque("Emitente").Value
    If rsCheque("Tipo do Movimento").Value = "1" Then
        Mid(x_linha, 125, 12) = "Combustível "
    ElseIf rsCheque("Tipo do Movimento").Value = "2" Then
        Mid(x_linha, 125, 12) = "Oleo/Diverso"
    ElseIf rsCheque("Tipo do Movimento").Value = "3" Then
        Mid(x_linha, 125, 12) = "Inclusão    "
    End If
    If txt_taxa_juros.Text <> "" Then
        Mid(x_linha, 125, 12) = "            "
        xValor = rsCheque("valor").Value * (fValidaValor(txt_taxa_juros.Text) / 30 * xDias) / 100
        i = Len(Format(xValor, "####,##0.00"))
        Mid(x_linha, 125 + 11 - i, i) = Format(xValor, "####,##0.00")
        lTotalJuros = lTotalJuros + xValor
    End If
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
    
    If rsCheque("Dados do Abastecimento").Value <> Empty Then
        BioImprime "@@Printer.FontName = Sans Serif 17cpi"
        BioImprime "@Printer.Print " & PreencheLinhaDadosAbastecimento(rsCheque("Dados do Abastecimento").Value)
        lLinha = lLinha + 1
    End If
        

End Sub
Private Function PreencheLinhaDadosAbastecimento(ByVal pDadosAbastecimento As String) As String
    Dim xLinha As String
    Dim xBico As String
    Dim xData As String
    Dim xHora As String
    Dim xValor As Currency
    Dim xTipoCombustivel As String
    
    PreencheLinhaDadosAbastecimento = Empty
    On Error GoTo FileError
        
        xBico = RetiraString(1, pDadosAbastecimento)
        xData = RetiraString(2, pDadosAbastecimento)
        xHora = RetiraString(3, pDadosAbastecimento)
        xValor = CCur(RetiraString(4, pDadosAbastecimento))
        xTipoCombustivel = RetiraString(5, pDadosAbastecimento)
        
        xLinha = "|            |       |           |          |     |               |            |                                          |             |"
        Mid(xLinha, 69, 61) = "Bico: " & xBico & " Dt:" & xData & " Hr:" & xHora & " Vl:" & FormatNumber(xValor, 2) & " Comb: " & xTipoCombustivel
        
        PreencheLinhaDadosAbastecimento = xLinha
        Exit Function
FileError:
    Exit Function
End Function
Private Sub ImpSubTotal()
    Dim x_linha As String
    Dim i As Integer
    If lSubTotal > 0 Then
        x_linha = "|            |       |           |          |     |               |            |                                          |             |"
        Mid(x_linha, 35, 10) = "*** TOTAL "
        i = Len(Format(lSubDias / lSubQtd, "#0.00"))
        Mid(x_linha, 46 + 5 - i, i) = Format(lSubDias / lSubQtd, "#0.00")
        i = Len(Format(lSubTotal, "###,###,##0.00"))
        Mid(x_linha, 52 + 14 - i, i) = Format(lSubTotal, "###,###,##0.00")
        Mid(x_linha, 82, 17) = "NÚMERO DE CHEQUES"
        i = Len(Format(lSubQtd, "####"))
        Mid(x_linha, 104 + 4 - i, i) = Format(lSubQtd, "####")
        Mid(x_linha, 108, 4) = " EM "
        Mid(x_linha, 112, 10) = Format(l_data, "dd/mm/yyyy")
        BioImprime "@@Printer.FontName = Sans Serif 17cpi"
        BioImprime "@Printer.Print " & x_linha
        lLinha = lLinha + 1
        lSubTotal = 0
        lSubQtd = 0
        lSubDias = 0
    End If
End Sub
Private Sub ImpTotal()
    Dim x_linha As String
    Dim i As Integer
    BioImprime "@Printer.Print " & "+------------+-------+-----------+----------+-----+---------------+------------+------------------------------------------+-------------+"
    x_linha = "|                                           |     |               |            |                                          |             |"
    Mid(x_linha, 35, 10) = "*** TOTAL "
    i = Len(Format(lTotalDias / lTotalQtd, "#0.00"))
    Mid(x_linha, 46 + 5 - i, i) = Format(lTotalDias / lTotalQtd, "#0.00")
    i = Len(Format(lTotal, "###,###,##0.00"))
    Mid(x_linha, 52 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
    Mid(x_linha, 82, 17) = "NÚMERO DE CHEQUES"
    i = Len(Format(lTotalQtd, "####"))
    Mid(x_linha, 104 + 4 - i, i) = Format(lTotalQtd, "####")
    If txt_taxa_juros.Text <> "" Then
        i = Len(Format(lTotalJuros, "####,##0.00"))
        Mid(x_linha, 125 + 11 - i, i) = Format(lTotalJuros, "####,##0.00")
    End If
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "+-------------------------------------------+-----+---------------+------------+------------------------------------------+-------------+"
    Mid(x_linha, 84, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & "  "
End Sub
Private Sub ImpCab()
    Dim xLinha As String
    Dim x_string_40 As String * 40
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
    x_string_40 = g_nome_empresa
    BioImprime "@Printer.Print " & "| " & x_string_40 & "                         Página, " & Format(lPagina, "000") & " |"
    BioImprime "@Printer.Print " & "| RELAÇÃO DE CHEQUE PRÉ-DATADO                             Goiânia, " & msk_data & " |"
    x_string_40 = Mid(cbo_tipo_movimento, 3, Len(cbo_tipo_movimento))
    BioImprime "@Printer.Print " & "| Tipo de Movimento.: " & x_string_40 & "                 |"
    BioImprime "@Printer.Print " & "| Referente a.: " & msk_data_i & " a " & msk_data_f & "       Período " & cbo_periodo_i & " ao " & cbo_periodo_f & "                   |"
    If txt_taxa_juros.Text <> "" Then
        xLinha = "| Taxa de Juros:       % ao Mês                                                |"
        i = Format(fValidaValor(txt_taxa_juros.Text), "##0.00")
        Mid(xLinha, 18 + 6 - i, i) = Format(fValidaValor(txt_taxa_juros.Text), "##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    xLinha = "+------------+-------+-----------+----------+-----+---------------+------------+------------------------------------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|   EMISSÃO  |PERIODO|N. DA CONTA| N.CHEQUE |PRAZO|VALOR DO CHEQUE| VENCIMENTO | NOME DO EMITENTE                         | TIPO   MOV. |"
    If txt_taxa_juros.Text <> "" Then
        Mid(xLinha, 124, 13) = " VALOR JUROS "
    End If
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+------------+-------+-----------+----------+-----+---------------+------------+------------------------------------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
End Sub
Private Sub cbo_periodo_f_GotFocus()
    SendMessageLong cbo_periodo_f.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub cbo_periodo_i_GotFocus()
    SendMessageLong cbo_periodo_i.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_f.ListIndex = cbo_periodo_i.ListIndex
        cbo_periodo_f.SetFocus
    End If
End Sub
Private Sub cbo_tipo_movimento_GotFocus()
    SendMessageLong cbo_tipo_movimento.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_tipo_movimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        optEmissao.SetFocus
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
    optEmissao.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        optEmissao.SetFocus
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
        MsgBox "Data final deve ser maior que a data inicial.", 64, "Atenção!"
        msk_data_f.SetFocus
    ElseIf optEmissao.Value = False And optVencimento.Value = False Then
        MsgBox "Escolha o tipo de emissão.", 64, "Atenção!"
        optEmissao.SetFocus
    ElseIf cbo_periodo_i.ListIndex = -1 Then
        MsgBox "Escolha o período inicial.", 64, "Atenção!"
        cbo_periodo_i.SetFocus
    ElseIf cbo_periodo_f.ListIndex = -1 Then
        MsgBox "Escolha o período final.", 64, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cbo_periodo_f.Text < cbo_periodo_i.Text Then
        MsgBox "O periodo final deve ser maior que " & Val(cbo_periodo_i) - 1 & ".", 64, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cbo_tipo_movimento.ListIndex = -1 Then
        MsgBox "Escolha o tipo de movimento.", 64, "Atenção!"
        cbo_tipo_movimento.SetFocus
    ElseIf chkChequeVista.Value = 0 And chkChequePredatado.Value = 0 Then
        MsgBox "Marque um tipo de cheque (Pre-Datado ou A Vista).", 64, "Atenção!"
        chkChequePredatado.SetFocus
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
        cbo_periodo_i.ListIndex = 0
        cbo_periodo_f.ListIndex = 0
        cbo_tipo_movimento.ListIndex = 0
        cbo_periodo_i.SetFocus
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
    PreencheCboPeriodo
    PreencheCboTipoMovimento
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
        optEmissao.SetFocus
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
Private Sub optEmissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_i.SetFocus
    End If
End Sub
Private Sub optVencimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_i.SetFocus
    End If
End Sub
Private Sub txt_taxa_juros_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_taxa_juros_LostFocus()
    txt_taxa_juros.Text = Format(txt_taxa_juros.Text, "###,##0.00")
End Sub
