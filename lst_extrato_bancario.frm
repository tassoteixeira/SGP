VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form lst_extrato_bancario 
   Caption         =   "Emissão do Extrato Bancário"
   ClientHeight    =   2850
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6945
   Icon            =   "lst_extrato_bancario.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_extrato_bancario.frx":030A
   ScaleHeight     =   2850
   ScaleWidth      =   6945
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_extrato_bancario.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Visualiza extrato bancário"
      Top             =   1860
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_extrato_bancario.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "extrato bancário"
      Top             =   1860
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_extrato_bancario.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1860
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6675
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_extrato_bancario.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_extrato_bancario.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   6060
         Picture         =   "lst_extrato_bancario.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4980
         TabIndex        =   10
         Top             =   1080
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
         TabIndex        =   7
         Top             =   1080
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
         TabIndex        =   4
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
      Begin MSAdodcLib.Adodc adodc_conta 
         Height          =   330
         Left            =   3510
         Top             =   240
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adodc_conta"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo dtcbo_conta 
         Bindings        =   "lst_extrato_bancario.frx":7F94
         Height          =   315
         Left            =   1650
         TabIndex        =   2
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_conta"
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         Caption         =   "&Conta"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   315
         Left            =   3960
         TabIndex        =   9
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   2100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_extrato_bancario"
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
Dim lTotCredito As Currency
Dim lTotDebito As Currency
Dim lNumeroConta As String
Dim lSQL As String
Private rsMovBancario As New adodb.Recordset
Private MovBancario As New cMovimentoBancario
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set rsMovBancario = Nothing
    Set MovBancario = Nothing
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lSaldo = 0
    lTotCredito = 0
    lTotDebito = 0
    lNumeroConta = ""
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    
    'Prepara SQL
    lSQL = "SELECT Data, [Numero do Movimento], Valor"
    lSQL = lSQL & ", [Debito ou Credito], [Numero do Documento], [Codigo do Historico]"
    lSQL = lSQL & ", Complemento, HistoricoPadrao.Nome"
    lSQL = lSQL & " FROM MovimentoBancario, HistoricoPadrao"
    lSQL = lSQL & " WHERE MovimentoBancario.Empresa = " & g_empresa
    lSQL = lSQL & " AND MovimentoBancario.[Codigo do Portador] = " & Chr(39) & dtcbo_conta.BoundText & Chr(39)
    lSQL = lSQL & " AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & " AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & " AND HistoricoPadrao.Codigo = MovimentoBancario.[Codigo do Historico]"
    lSQL = lSQL & " ORDER BY Data, [Numero do Movimento]"
    
    'Abre RecordSet
    Set rsMovBancario = New adodb.Recordset
    Set rsMovBancario = Conectar.RsConexao(lSQL)
    
    'Verifica movimento
    If rsMovBancario.RecordCount > 0 Then
        ImpDados
    End If
    If rsMovBancario.State = 1 Then
        rsMovBancario.Close
    End If
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    Dim xLinha As String
    'loop movimento bancário
    Do Until rsMovBancario.EOF
        If lPagina = 0 Then
            ImpCab
        End If
        If lLinha >= 57 Then
            xLinha = "+------------+-----------------+-----------------+-----------------+---+----------------------------------------------------------------+"
            Mid(xLinha, 76, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & xLinha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        ImpDet
        rsMovBancario.MoveNext
    Loop
    If lPagina > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório do Extrato Bancário|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub ImpDet()
    Dim xLinha As String
    Dim i As Integer
    '''               10        20        30        40        50        60        70        80        90       100       110       120       130
    '''       12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    xLinha = "|            |                 |                 |                 |   |                                                     |          |"
    Mid(xLinha, 3, 10) = Format(rsMovBancario("Data").Value, "dd/mm/yyyy")
    If rsMovBancario("Debito ou Credito").Value = "C" Then
        i = Len(Format(rsMovBancario("Valor").Value, "##,###,##0.00"))
        Mid(xLinha, 18 + 13 - i, i) = Format(rsMovBancario("Valor").Value, "##,###,##0.00")
        lTotCredito = lTotCredito + rsMovBancario("Valor").Value
        lSaldo = lSaldo + rsMovBancario("Valor").Value
    Else
        i = Len(Format(rsMovBancario("Valor").Value, "##,###,##0.00"))
        Mid(xLinha, 36 + 13 - i, i) = Format(rsMovBancario("Valor").Value, "##,###,##0.00")
        lTotDebito = lTotDebito + rsMovBancario("Valor").Value
        lSaldo = lSaldo - rsMovBancario("Valor").Value
    End If
    i = Len(Format(lSaldo, "##,###,##0.00"))
    Mid(xLinha, 54 + 13 - i, i) = Format(lSaldo, "##,###,##0.00")
    i = Len(Format(rsMovBancario("Codigo do Historico").Value, "###"))
    Mid(xLinha, 69 + 3 - i, i) = Format(rsMovBancario("Codigo do Historico").Value, "###")
    Mid(xLinha, 73, 51) = Trim(rsMovBancario("Nome").Value) & " " & rsMovBancario("Complemento").Value
    Mid(xLinha, 127, 10) = rsMovBancario("Numero do Documento").Value
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    BioImprime "@Printer.Print " & "+------------+-----------------+-----------------+-----------------+---+-----------------------------------------------------+----------+"
    '''               10        20        30        40        50        60        70        80        90       100       110       120       130
    '''       12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    xLinha = "|            |                 |                 |                 |                                                                    |"
    Mid(xLinha, 4, 10) = "*** TOTAL "
    i = Len(Format(lTotCredito, "##,###,##0.00"))
    Mid(xLinha, 18 + 13 - i, i) = Format(lTotCredito, "##,###,##0.00")
    i = Len(Format(lTotDebito, "##,###,##0.00"))
    Mid(xLinha, 36 + 13 - i, i) = Format(lTotDebito, "##,###,##0.00")
    i = Len(Format(lSaldo, "##,###,##0.00"))
    Mid(xLinha, 54 + 13 - i, i) = Format(lSaldo, "##,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+------------+-----------------+-----------------+-----------------+--------------------------------------------------------------------+"
    Mid(xLinha, 72, 22) = " Cerrado Informática. "
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
    xLinha = "| EMISSAO DO EXTRATO BANCARIO                               CIDADE, __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| Referente a.: __/__/____ a __/__/____                                        |"
    Mid(xLinha, 17, 10) = msk_data_i.Text
    Mid(xLinha, 30, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| Conta:           -                                                           |"
    Mid(xLinha, 10, 10) = dtcbo_conta.BoundText
    Mid(xLinha, 22, 40) = dtcbo_conta.Text
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    '''                                     10        20        30        40        50        60        70        80        90       100       110       120       130
    '''                             12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    BioImprime "@Printer.Print " & "+------------+-----------------+-----------------+-----------------+---+-----------------------------------------------------+----------+"
    BioImprime "@Printer.Print " & "|DATA DO MOV.|     CRÉDITO     |     DÉBITO      |      SALDO      |COD|HISTÓRICO / COMPLEMENTO                              |N.DOCUMENT|"
    BioImprime "@Printer.Print " & "+------------+-----------------+-----------------+-----------------+---+-----------------------------------------------------+----------+"
    lNumeroConta = ""
    If lPagina = 1 Then
        xLinha = "|            |                 |                 |                 |   |SALDO ANTERIOR                                       |          |"
        lSaldo = MovBancario.BuscaSaldoAnterior(g_empresa, dtcbo_conta.BoundText, CDate(msk_data_i.Text))
        i = Len(Format(lSaldo, "##,###,##0.00"))
        Mid(xLinha, 54 + 13 - i, i) = Format(lSaldo, "##,###,##0.00")
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
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
Private Sub dtcbo_conta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        If MovBancario.LocalizarUltimo(g_empresa, dtcbo_conta.BoundText) Then
            'AtivaBotoes
            'AtualTela
            msk_data.SetFocus
        Else
            'DesativaBotoes
            'cmd_novo.Enabled = True
            cmd_sair.Enabled = True
            'LimpaTela
            MsgBox "Não há registros nesta conta.", vbInformation, "Erro de Verificação!"
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        dtcbo_conta.SetFocus
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
    
'    adodc_conta.ConnectionString = gConnectionString
'    adodc_conta.RecordSource = "SELECT Codigo, Nome FROM Conta_Bancaria WHERE Empresa = " & g_empresa & " ORDER BY Nome"
'    adodc_conta.Refresh
    Set adodc_conta.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM PortadorFinanceiro WHERE Empresa = " & g_empresa & " ORDER BY Nome")
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
