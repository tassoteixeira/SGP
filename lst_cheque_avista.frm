VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_cheque_avista 
   Caption         =   "Emissão dos Cheques à Vista"
   ClientHeight    =   3075
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_cheque_avista.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_cheque_avista.frx":030A
   ScaleHeight     =   3075
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_cheque_avista.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2100
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_cheque_avista.frx":19E2
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprime cheques à vista."
      Top             =   2100
      Width           =   795
   End
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_cheque_avista.frx":2FEC
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Visualiza cheques à vista."
      Top             =   2100
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1875
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_cheque_avista.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_cheque_avista.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "lst_cheque_avista.frx":6CBA
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
         TabIndex        =   15
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox cbo_periodo_f 
         Height          =   315
         Left            =   5940
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1020
         Width           =   495
      End
      Begin VB.ComboBox cbo_periodo_i 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1020
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4800
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
         Caption         =   "Período &final"
         Height          =   315
         Left            =   3840
         TabIndex        =   12
         Top             =   1020
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "&Tipo de Movimento"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Período inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1020
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
         Width           =   915
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
      Top             =   2340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_cheque_avista"
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
Dim lSQl As String

Private rsCheque As New adodb.Recordset
Private rsChequeVista As New adodb.Recordset
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
    cbo_tipo_movimento.AddItem "1 Caixa de Combustíveis"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 1
    cbo_tipo_movimento.AddItem "2 Caixa de Óleo/Diversos"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 2
    cbo_tipo_movimento.AddItem "3 Cheque Inclusão"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 3
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lSubTotal = 0
    lTotal = 0
    lSubQtd = 0
    lTotalQtd = 0
End Sub
Private Sub Relatorio()
    Dim lok As Boolean
    lok = False
    ZeraVariaveis
    
    'Verifica Movimento_Cheque
    'Prepara SQL
    lSQl = ""
    lSQl = lSQl & "SELECT [Data do Vencimento], [Data de Emissao], Valor"
    lSQl = lSQl & "  FROM Movimento_Cheque"
    lSQl = lSQl & " WHERE Empresa = " & g_empresa
    lSQl = lSQl & "   AND [Data do Vencimento] >= " & preparaData(msk_data_i.Text)
    lSQl = lSQl & "   AND [Data do Vencimento] <= " & preparaData(msk_data_f.Text)
    lSQl = lSQl & "   AND Periodo >= " & preparaTexto(cbo_periodo_i.Text)
    lSQl = lSQl & "   AND Periodo <= " & preparaTexto(cbo_periodo_f.Text)
    If Val(cbo_tipo_movimento.Text) > 0 Then
        lSQl = lSQl & "   AND [Tipo do Movimento] = " & preparaTexto(Val(cbo_tipo_movimento.Text))
    End If
    lSQl = lSQl & " ORDER BY [Data do Vencimento], [Data de Emissao], Periodo, [Tipo do Movimento], [Ordem da Digitacao], [Numero da Conta], [Numero do Cheque]"
    'Abre RecordSet
    Set rsCheque = New adodb.Recordset
    Set rsCheque = Conectar.RsConexao(lSQl)
    If rsCheque.RecordCount > 0 Then
        lok = True
    End If
    
    
    
    
    'Verifica Movimento_Cheque_Avista
    'Prepara SQL
    lSQl = ""
    lSQl = lSQl & "SELECT [Data de Emissao], Valor"
    lSQl = lSQl & "  FROM Movimento_Cheque_Avista"
    lSQl = lSQl & " WHERE Empresa = " & g_empresa
    lSQl = lSQl & "   AND [Data de Emissao] >= " & preparaData(msk_data_i.Text)
    lSQl = lSQl & "   AND [Data de Emissao] <= " & preparaData(msk_data_f.Text)
    lSQl = lSQl & "   AND Periodo >= " & preparaTexto(cbo_periodo_i.Text)
    lSQl = lSQl & "   AND Periodo <= " & preparaTexto(cbo_periodo_f.Text)
    If Val(cbo_tipo_movimento.Text) > 0 Then
        lSQl = lSQl & "   AND [Tipo do Movimento] = " & preparaTexto(Val(cbo_tipo_movimento.Text))
    End If
    lSQl = lSQl & " ORDER BY [Data de Emissao], Periodo, [Tipo do Movimento], [Ordem da Digitacao]"
    'Abre RecordSet
    Set rsChequeVista = New adodb.Recordset
    Set rsChequeVista = Conectar.RsConexao(lSQl)
    If rsChequeVista.RecordCount > 0 Then
        lok = True
    End If
    
    
    
    If lok Then
        ImpDados
    End If
    
    
    If rsCheque.State = 1 Then
        rsCheque.Close
    End If
    If rsChequeVista.State = 1 Then
        rsChequeVista.Close
    End If
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    Dim x_linha As String
    Call LoopChequeAvista("N")
    Call ImpSubTotal("*** TOTAL CHEQUE À VISTA")
    Call LoopChequeAvista("S")
    Call ImpSubTotal("*** TOTAL CHEQUE INCLUSÃO")
    If g_empresa <> 2 Then
        LoopCheque
        Call ImpSubTotal("*** TOTAL CHEQUE PRÉ-DATADO")
    End If
    If lTotal > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Cheque à Vista|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub LoopCheque()
    Dim x_linha As String
    
    
    'loop movimento de cheques
    rsCheque.MoveFirst
    Do Until rsCheque.EOF
        If lPagina = 0 Then
            ImpCab
            l_data = rsCheque("Data do Vencimento").Value
        End If
        If lLinha >= 60 Then
            x_linha = String(80, "-")
            Mid(x_linha, 1, 1) = "+"
            Mid(x_linha, 31, 1) = "+"
            Mid(x_linha, 49, 1) = "+"
            Mid(x_linha, 80, 1) = "+"
            BioImprime "@Printer.Print " & x_linha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        If l_data <> rsCheque("Data do Vencimento").Value Then
            Call ImpSubTotal("*** TOTAL CHEQUE PRÉ-DATADO")
            l_data = rsCheque("Data do Vencimento").Value
        End If
        Call ImpDet(rsCheque("Valor").Value)
        lSubTotal = lSubTotal + rsCheque("Valor").Value
        lTotal = lTotal + rsCheque("Valor").Value
        lSubQtd = lSubQtd + 1
        lTotalQtd = lTotalQtd + 1
        rsCheque.MoveNext
    Loop
End Sub
Private Sub LoopChequeAvista(x_inclusao As String)
    'loop movimento de cheques à vista
    Dim x_linha As String
    
    rsChequeVista.MoveFirst
    Do Until rsChequeVista.EOF
        If (x_inclusao = "S" And rsChequeVista("Tipo do Movimento").Value = 3) Or x_inclusao = "N" And rsChequeVista("Tipo do Movimento").Value <> 3 Then
            If lPagina = 0 Then
                ImpCab
                l_data = rsChequeVista("Data de Emissao").Value
            End If
            If lLinha >= 60 Then
                x_linha = String(80, "-")
                Mid(x_linha, 1, 1) = "+"
                Mid(x_linha, 31, 1) = "+"
                Mid(x_linha, 49, 1) = "+"
                Mid(x_linha, 80, 1) = "+"
                BioImprime "@Printer.Print " & x_linha
                BioImprime "@@Printer.NewPage"
                ImpCab
            End If
            Call ImpDet(rsChequeVista("Valor").Value)
            lSubTotal = lSubTotal + rsChequeVista("Valor").Value
            lTotal = lTotal + rsChequeVista("Valor").Value
            lSubQtd = lSubQtd + 1
            lTotalQtd = lTotalQtd + 1
        End If
        rsChequeVista.MoveNext
    Loop
End Sub
Private Sub ImpDet(x_valor As Currency)
    Dim x_linha As String
    Dim i As Integer
    x_linha = Space(80)
    
    Mid(x_linha, 1, 1) = "|"
    Mid(x_linha, 31, 1) = "|"
    i = Len(Format(x_valor, "###,###,##0.00"))
    Mid(x_linha, 33 + 14 - i, i) = Format(x_valor, "###,###,##0.00")
    Mid(x_linha, 49, 1) = "|"
    Mid(x_linha, 80, 1) = "|"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpSubTotal(x_mensagem As String)
    Dim y_local As Single
    Dim x_linha As String * 80
    Dim i As Integer
    If lSubTotal > 0 Then
        x_linha = String(80, " ")
        Mid(x_linha, 1, 1) = "|"
        Mid(x_linha, 31, 19) = "+-----------------+"
        Mid(x_linha, 80, 1) = "|"
        BioImprime "@Printer.Print " & x_linha
        x_linha = Space(80)
        Mid(x_linha, 1, 1) = "|"
        i = Len(Trim(x_mensagem))
        Mid(x_linha, 1 + 29 - i, i) = Trim(x_mensagem)
        Mid(x_linha, 31, 1) = "|"
        Mid(x_linha, 49, 1) = "|"
        i = Len(Format(lSubTotal, "###,###,##0.00"))
        Mid(x_linha, 33 + 14 - i, i) = Format(lSubTotal, "###,###,##0.00")
        i = Len(Format(lSubQtd, "####"))
        Mid(x_linha, 51 + 4 - i, i) = Format(lSubQtd, "####")
        Mid(x_linha, 80, 1) = "|"
        BioImprime "@@Printer.FontName = Draft 10cpi"
        BioImprime "@@Printer.FontBold = True"
        BioImprime "@@y_local = Printer.CurrentY"
        BioImprime "@@Printer.FontBold = True"
        BioImprime "@Printer.Print " & x_linha
'        Printer.CurrentY = y_local - 0.01
'        Printer.Print x_linha
        BioImprime "@@Printer.CurrentY = y_local"
        BioImprime "@@Printer.Print " & "  "
        BioImprime "@@Printer.FontBold = False"
        x_linha = String(80, "-")
        Mid(x_linha, 1, 1) = "+"
        Mid(x_linha, 31, 1) = "+"
        Mid(x_linha, 49, 1) = "+"
        Mid(x_linha, 80, 1) = "+"
        BioImprime "@Printer.Print " & x_linha
        lLinha = lLinha + 3
        lSubTotal = 0
        lSubQtd = 0
    End If
End Sub
Private Sub ImpTotal()
    Dim y_local As Single
    Dim x_linha As String * 80
    Dim i As Integer
    x_linha = Space(80)
    Mid(x_linha, 1, 1) = "|"
    Mid(x_linha, 15, 16) = "*** TOTAL GERAL "
    Mid(x_linha, 31, 1) = "|"
    Mid(x_linha, 49, 1) = "|"
    i = Len(Format(lTotal, "###,###,##0.00"))
    Mid(x_linha, 33 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
    i = Len(Format(lTotalQtd, "####"))
    Mid(x_linha, 51 + 4 - i, i) = Format(lTotalQtd, "####")
    Mid(x_linha, 80, 1) = "|"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@@y_local = Printer.CurrentY"
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
'    Printer.CurrentY = y_local - 0.01
'    Printer.Print x_linha
    BioImprime "@@Printer.CurrentY = y_local"
    BioImprime "@@Printer.Print " & "  "
    BioImprime "@@Printer.FontBold = False"
    x_linha = String(80, "-")
    Mid(x_linha, 1, 1) = "+"
    Mid(x_linha, 31, 1) = "+"
    Mid(x_linha, 49, 1) = "+"
    Mid(x_linha, 52, 22) = " Cerrado Informática. "
    Mid(x_linha, 80, 1) = "+"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpCab()
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
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & "| " & x_string_40 & "                         Página, " & Format(lPagina, "000") & " |"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "| RELAÇÃO DOS CHEQUES PARA DEPÓSITO                        Goiânia, " & msk_data & " |"
    x_string_40 = Mid(cbo_tipo_movimento, 3, Len(cbo_tipo_movimento))
    BioImprime "@Printer.Print " & "| Tipo de Movimento.: " & x_string_40 & "                 |"
    BioImprime "@Printer.Print " & "| Referente a.: " & msk_data_i & " a " & msk_data_f & "       Período " & cbo_periodo_i & " ao " & cbo_periodo_f & "                   |"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & "+-----------------------------+-----------------+------------------------------+"
    BioImprime "@Printer.Print " & "|                             | VALOR DO CHEQUE |                              |"
    BioImprime "@Printer.Print " & "+-----------------------------+-----------------+------------------------------+"
End Sub
Private Sub cbo_periodo_f_GotFocus()
    SendMessageLong cbo_periodo_f.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo_movimento.SetFocus
    End If
End Sub
Private Sub cbo_periodo_i_GotFocus()
    SendMessageLong cbo_periodo_i.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_f.SetFocus
    End If
End Sub
Private Sub cbo_tipo_movimento_GotFocus()
    SendMessageLong cbo_tipo_movimento.hwnd, CB_SHOWDROPDOWN, True, 0
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
        cbo_periodo_i.SetFocus
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
    cbo_periodo_i.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cbo_periodo_i.SetFocus
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
        msk_data_i.Text = Format(g_data_def, "dd/mm/yyyy")
        If Format(g_data_def, "ddd") = "Seg" Then
            msk_data_i.Text = Format(g_data_def - 2, "dd/mm/yyyy")
        End If
        msk_data_f.Text = Format(g_data_def, "dd/mm/yyyy")
        cbo_periodo_i.ListIndex = 0
        cbo_periodo_f.ListIndex = 3
        cbo_tipo_movimento.ListIndex = 0
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
    PreencheCboPeriodo
    PreencheCboTipoMovimento
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_i.SetFocus
    End If
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
