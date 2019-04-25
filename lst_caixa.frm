VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form lst_caixa 
   Caption         =   "Emissão do Caixa"
   ClientHeight    =   2310
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_caixa.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_caixa.frx":030A
   ScaleHeight     =   2310
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_caixa.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Visualiza cheque devolvido."
      Top             =   1380
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_caixa.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprime caixa."
      Top             =   1380
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_caixa.frx":3074
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
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_caixa.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_caixa.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "lst_caixa.frx":6CBA
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
Dim lSaldo As Currency
Dim lTotalCredito As Currency
Dim lTotalDebito As Currency
Dim lNumeroConta As String
Dim lSQL As String
Private rsMovCaixa As New adodb.Recordset
Private Sub Finaliza()
    Set rsMovCaixa = Nothing
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lSaldo = 0
    lTotalCredito = 0
    lTotalDebito = 0
    lNumeroConta = ""
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    
    'Prepara SQL
    lSQL = "SELECT Data, [Numero do Movimento], Valor"
    lSQL = lSQL & ", [Debito ou Credito], [Numero do Documento], [Codigo do Historico]"
    lSQL = lSQL & ", Complemento, [Numero da Conta], [Tipo do Movimento], Plano_Conta.Nome, HistoricoPadrao.Nome as NomeHistorico"
    lSQL = lSQL & " FROM Movimento_Caixa, Plano_Conta, HistoricoPadrao"
    lSQL = lSQL & " WHERE Movimento_Caixa.Empresa = " & g_empresa
    lSQL = lSQL & " AND Data >= " & Chr(35) & Format(CDate(msk_data_i.Text), "mm/dd/yyyy") & Chr(35)
    lSQL = lSQL & " AND Data <= " & Chr(35) & Format(CDate(msk_data_f.Text), "mm/dd/yyyy") & Chr(35)
    lSQL = lSQL & " AND Plano_Conta.Codigo = Movimento_Caixa.[Numero da Conta]"
    lSQL = lSQL & " AND HistoricoPadrao.Codigo = Movimento_Caixa.[Codigo do Historico]"
    lSQL = lSQL & " ORDER BY [Numero do Documento], Data, [Numero do Movimento]"
    
    'Abre RecordSet
    Set rsMovCaixa = New adodb.Recordset
    Set rsMovCaixa = Conectar.RsConexao(lSQL)
    
    'Verifica movimento
    If rsMovCaixa.RecordCount > 0 Then
        ImpDados
    End If
    If rsMovCaixa.State = 1 Then
        rsMovCaixa.Close
    End If
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    Dim xLinha As String
    'loop movimento de caixa
    Do Until rsMovCaixa.EOF
        If lPagina = 0 Then
            ImpCab
        End If
        If lLinha >= 57 Then
            xLinha = "+------------+-----------+-----------------+-----------------+-----------------+"
            Mid(xLinha, 84, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & xLinha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        ImpDet
        rsMovCaixa.MoveNext
    Loop
    If lPagina > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório do Movimento do Caixa|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub ImpConta()
    Dim xLinha As String
    Dim i As Integer
    BioImprime "@Printer.Print " & "+------------+-----------+-----------------+-----------------+-----------------+"
    xLinha = "| Conta Contábil ...:               -                                          |"
    Mid(xLinha, 23, 13) = fMascaraContaContabil(rsMovCaixa("Numero da Conta").Value)
    Mid(xLinha, 39, 40) = rsMovCaixa("Nome").Value
    BioImprime "@Printer.Print " & xLinha
    lNumeroConta = rsMovCaixa("Numero da Conta").Value
    lSaldo = 0
    lLinha = lLinha + 2
End Sub
Private Sub ImpDet()
    Dim xLinha As String
    Dim i As Integer
    '''               10        20        30        40        50        60        70        80        90       100       110       120       130
    '''       12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    If lNumeroConta <> rsMovCaixa("Numero da Conta").Value Then
        Call ImpConta
    End If
    xLinha = "|            |           |                 |                 |                 |"
    Mid(xLinha, 3, 10) = Format(rsMovCaixa("Data").Value, "dd/mm/yyyy")
    i = Len(Format(rsMovCaixa("Numero do Movimento").Value, "###,##0"))
    Mid(xLinha, 18 + 7 - i, i) = Format(rsMovCaixa("Numero do Movimento").Value, "###,##0")
    If rsMovCaixa("Debito ou Credito").Value = "C" Then
        i = Len(Format(rsMovCaixa("Valor").Value, "##,###,##0.00"))
        Mid(xLinha, 30 + 13 - i, i) = Format(rsMovCaixa("Valor").Value, "##,###,##0.00")
        lTotalCredito = lTotalCredito + rsMovCaixa("Valor").Value
        lSaldo = lSaldo + rsMovCaixa("Valor").Value
    Else
        i = Len(Format(rsMovCaixa("Valor").Value, "##,###,##0.00"))
        Mid(xLinha, 48 + 13 - i, i) = Format(rsMovCaixa("Valor").Value, "##,###,##0.00")
        lTotalDebito = lTotalDebito + rsMovCaixa("Valor").Value
        lSaldo = lSaldo - rsMovCaixa("Valor").Value
    End If
    i = Len(Format(lSaldo, "##,###,##0.00"))
    Mid(xLinha, 66 + 13 - i, i) = Format(lSaldo, "##,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
    
    '''               10        20        30        40        50        60        70        80        90       100       110       120       130
    '''       12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    xLinha = "|   |                                                               |          |"
    i = Len(Format(rsMovCaixa("Codigo do Historico").Value, "###"))
    Mid(xLinha, 2 + 3 - i, i) = Format(rsMovCaixa("Codigo do Historico").Value, "###")
    Mid(xLinha, 6, 63) = rsMovCaixa("NomeHistorico").Value & " " & rsMovCaixa("Complemento").Value
    Mid(xLinha, 70, 10) = rsMovCaixa("Numero do Documento").Value
    BioImprime "@Printer.Print " & xLinha
  
    
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    BioImprime "@Printer.Print " & "+------------+-----------+-----------------+-----------------+-----------------+"
    '''               10        20        30        40        50        60        70        80        90       100       110       120       130
    '''       12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    xLinha = "|                        |                 |                 |                 |"
    Mid(xLinha, 16, 10) = "*** TOTAL "
    
    i = Len(Format(lTotalCredito, "##,###,##0.00"))
    Mid(xLinha, 30 + 13 - i, i) = Format(lTotalCredito, "##,###,##0.00")
    i = Len(Format(lTotalDebito, "##,###,##0.00"))
    Mid(xLinha, 48 + 13 - i, i) = Format(lTotalDebito, "##,###,##0.00")
    
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+------------+-----------+-----------------+-----------------+-----------------+"
    Mid(xLinha, 3, 22) = " Cerrado Informática. "
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
    xLinha = "| RELAÇÃO DO MOVIMENTO DO CAIXA                             CIDADE, __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| Referente a.: __/__/____ a __/__/____                                        |"
    Mid(xLinha, 17, 10) = msk_data_i.Text
    Mid(xLinha, 30, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.FontBold = False"
    '''                                     10        20        30        40        50        60        70        80        90       100       110       120       130
    '''                             12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    BioImprime "@Printer.Print " & "+------------+-----------+-----------------+-----------------+-----------------+"
    BioImprime "@Printer.Print " & "|DATA DO MOV.|N.MOVIMENTO|     CRÉDITO     |     DÉBITO      |      SALDO      |"
    BioImprime "@Printer.Print " & "|COD|HISTÓRICO / COMPLEMENTO                                        |N.DOCUMENT|"
    lNumeroConta = ""
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
            Relatorio
        End If
    End If
End Sub
Private Sub Form_Activate()
    If Not IsDate(msk_data) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def - 1, "dd/mm/yyyy")
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
    'Set tbl_movimento_caixa = bd_sgle.OpenTable("Movimento_Caixa")
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
