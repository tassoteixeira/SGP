VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_cliente 
   Caption         =   "Emissão dos Clientes"
   ClientHeight    =   4515
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   7290
   Icon            =   "lst_cliente.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_cliente.frx":030A
   ScaleHeight     =   4515
   ScaleWidth      =   7290
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1320
      Picture         =   "lst_cliente.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Visualiza os clientes em ordem alfabética."
      Top             =   3540
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3240
      Picture         =   "lst_cliente.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Imprime os clientes em ordem alfabética."
      Top             =   3540
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5160
      Picture         =   "lst_cliente.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3540
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   3315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7035
      Begin VB.CheckBox chk_SomenteNota 
         Height          =   255
         Left            =   6660
         TabIndex        =   7
         Top             =   660
         Width           =   195
      End
      Begin VB.OptionButton optRel2 
         Caption         =   "Usar Modelo de Relatório 2"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   2460
         Width           =   2835
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   2580
         Width           =   6735
         Begin VB.OptionButton optResumido 
            Caption         =   "Resumido (Nome, cpf, cnpj)"
            Height          =   255
            Left            =   3720
            TabIndex        =   17
            Top             =   240
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.OptionButton optCompleto 
            Caption         =   "Completo"
            Height          =   255
            Left            =   1080
            TabIndex        =   16
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.OptionButton optRel1 
         Caption         =   "Usar Modelo de Relatório 1"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1500
         Value           =   -1  'True
         Width           =   2835
      End
      Begin VB.Frame frmModelo1 
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   1620
         Width           =   6735
         Begin VB.OptionButton optTipoVencimento 
            Caption         =   "Tipo de Vencimento"
            Height          =   255
            Left            =   1080
            TabIndex        =   12
            Top             =   240
            Width           =   2295
         End
         Begin VB.OptionButton optFormaPagamento 
            Caption         =   "Forma de Pagamento"
            Height          =   255
            Left            =   3720
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   3060
         Picture         =   "lst_cliente.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbo_vencimento 
         Height          =   315
         ItemData        =   "lst_cliente.frx":59E0
         Left            =   1980
         List            =   "lst_cliente.frx":59E2
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1020
         Width           =   4935
      End
      Begin VB.CheckBox chk_geral 
         Height          =   255
         Left            =   1980
         TabIndex        =   5
         Top             =   660
         Width           =   195
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   315
         Left            =   1980
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
      Begin VB.Label Label3 
         Caption         =   "Somente Nota de Abastecimento"
         Height          =   315
         Index           =   0
         Left            =   3840
         TabIndex        =   6
         Top             =   660
         Width           =   2715
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de &vencimento"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   1020
         Width           =   1755
      End
      Begin VB.Label Label3 
         Caption         =   "Imprime &todas empresas"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1755
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1755
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport cr_relato 
      Left            =   720
      Top             =   3780
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\VB5\Sgp\Data\lst_cliente.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
End
Attribute VB_Name = "lst_cliente"
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
Dim lSQL As String

Private Vencimento As New cVencimento
Private rsCliente As New adodb.Recordset

Private Sub AtivaBotoes(ByVal pAtiva As Boolean)
    cmd_visualizar.Enabled = pAtiva
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
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
End Sub
Private Sub PreencheCboVencimento()
    Dim rsTabela As New adodb.Recordset
    cbo_vencimento.Clear
    cbo_vencimento.AddItem "Todos os Vencimentos"
    cbo_vencimento.ItemData(cbo_vencimento.NewIndex) = 0
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Codigo, [Dia 1 Inicial], [Dia 1 Final], [Dia 1 Vencimento], [Dia 2 Inicial], [Dia 2 Final], [Dia 2 Vencimento]"
    lSQL = lSQL & "  FROM Vencimento"
    lSQL = lSQL & " ORDER BY Codigo"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            cbo_vencimento.AddItem rsTabela("Codigo").Value & " - Dia " & rsTabela("Dia 1 Inicial").Value & " ao " & rsTabela("Dia 1 Final").Value & "  Venc. dia " & rsTabela("Dia 1 Vencimento").Value & " e Dia " & rsTabela("Dia 2 Inicial").Value & " ao " & rsTabela("Dia 2 Final").Value & "  Venc. dia " & rsTabela("Dia 2 Vencimento").Value
            cbo_vencimento.ItemData(cbo_vencimento.NewIndex) = rsTabela("Codigo").Value
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
    lSQL = "SELECT Codigo, CEP, Telefone, [Codigo do Vencimento],"
    lSQL = lSQL & " [Razao Social], CPF, Endereco, Identidade, [Orgao Emissor],"
    lSQL = lSQL & " Bairro, [Inscricao Estadual], Cidade, UF, CGC"
    lSQL = lSQL & " FROM Cliente"
    lSQL = lSQL & " WHERE Inativo = " & preparaBooleano(False)
    If chk_geral.Value = 0 Then
        lSQL = lSQL & " AND Empresa = " & g_empresa
    End If
    If cbo_vencimento.ListIndex > 0 Then
        lSQL = lSQL & " AND [Codigo do Vencimento] = " & cbo_vencimento.ListIndex
    End If
    If chk_SomenteNota.Value = 1 Then
        lSQL = lSQL & " AND [Gera Nota de Abastecimento] = 1"
    End If
    lSQL = lSQL & " ORDER BY [Razao Social]"
    
    'Abre RecordSet
    Set rsCliente = New adodb.Recordset
    Set rsCliente = Conectar.RsConexao(lSQL)
    
    'Verifica movimento
    If rsCliente.RecordCount > 0 Then
        ImpDados
    End If
    If rsCliente.State = 1 Then
        rsCliente.Close
    End If
    BioFechaImprime
    g_string = lLocal & lNomeArquivo & "|@|Relatório de Clientes|@|"
    frm_preview.Show 1
End Sub
Private Sub RelatorioCR()
    Dim l_sql As String
    If ValidaCampos Then
        If bdAccess Then
            If optTipoVencimento.Value = True Then
                cr_relato.ReportFileName = "\VB5\Sgp\Data\lst_cliente.rpt"
            Else
                cr_relato.ReportFileName = "\VB5\Sgp\Data\lst_cliente2.rpt"
            End If
        ElseIf bdSqlServer Then
            If optTipoVencimento.Value = True Then
                cr_relato.ReportFileName = "\VB5\Sgp\Data\SQL_lst_cliente.rpt"
            Else
                cr_relato.ReportFileName = "\VB5\Sgp\Data\SQL_lst_cliente2.rpt"
            End If
        End If
        cr_relato.SortFields(0) = "+{Cliente.Razao Social}"
        cr_relato.Formulas(0) = "f_empresa = BeforeReadingRecords;""" & g_nome_empresa & """"
        cr_relato.Formulas(1) = "f_data_hoje =  BeforeReadingRecords;""" & msk_data.Text & """"
        cr_relato.Formulas(2) = "f_vencimento = BeforeReadingRecords;""" & cbo_vencimento.Text & """"
        If bdAccess Then
            l_sql = "{Cliente.Inativo} = False"
        ElseIf bdSqlServer Then
            l_sql = "{Cliente.Inativo} = 0"
        End If
        If chk_geral.Value = 0 Then
            l_sql = l_sql & " And {Cliente.Empresa} = " & g_empresa
        End If
'        If chk_SomenteNota.Value = 1 Then
'            l_sql = l_sql & " And {Cliente.Gera Nota de Abastecimento} = 1"
'        End If
        If cbo_vencimento.ListIndex > 0 Then
            l_sql = l_sql & " And {Cliente.Codigo do Vencimento} = " & cbo_vencimento.ItemData(cbo_vencimento.ListIndex)
        End If
        cr_relato.Destination = 0
        cr_relato.SelectionFormula = l_sql
        If bdSqlServer Then
            cr_relato.Connect = "DSN=sgp_data;UID=sa;PWD=" & gSenhaBD
            cr_relato.Password = gSenhaBD
        End If
        cr_relato.Action = 1
    End If
End Sub
Private Sub ImpDados()
    Dim x_linha As String
    'loop cliente
    Do Until rsCliente.EOF
        If lPagina = 0 Then
            ImpCab
        End If
        If lLinha >= 55 Then
            If optRel2.Value = True Then
                If Me.optResumido.Value = True Then
                    x_linha = "+---------+------------------------------------------+-----------------+---------------------+------------------------------------------+"
                    Mid(x_linha, 14, 22) = " Cerrado Informática. "
                Else
                    x_linha = "+-------------------------------------------------+-------------------------------+--------------------------+--------------------------+"
                    Mid(x_linha, 4, 22) = " Cerrado Informática. "
                End If
            End If
            BioImprime "@Printer.Print " & x_linha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        If optCompleto.Value = True Then
            ImpDetCompleto
        Else
            ImpDetResumido
        End If
        rsCliente.MoveNext
    Loop
    ImpTotal
    Printer.EndDoc
End Sub
Private Sub ImpDetCompleto()
    Dim x_linha As String
    Dim i As Integer
    
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    x_linha = "+-------------------------------------------------+-------------------------------+--------------------------+--------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|CODIGO...:                                       |C.E.P....:       -             |FONE.: (   )    -         |DIA: ** A **    VENC.: ** |"
    Mid(x_linha, 13, 6) = Format(rsCliente("Codigo").Value, "00,000")
    Mid(x_linha, 63, 6) = Format(Mid(rsCliente("CEP").Value, 1, 5), "00,000")
    Mid(x_linha, 70, 4) = Mid(rsCliente("CEP").Value, 6, 3)
    Mid(x_linha, 92, 11) = fMascaraTelefone(rsCliente("Telefone").Value)
    If Vencimento.LocalizarCodigo(rsCliente("Codigo do Vencimento").Value) Then
        Mid(x_linha, 116, 2) = Format(Vencimento.Dia1Inicial, "00")
        Mid(x_linha, 121, 2) = Format(Vencimento.Dia1Final, "00")
        Mid(x_linha, 134, 2) = Format(Vencimento.Dia1Vencimento, "00")
    End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|NOME.....:                                       |CPF......:                     |                          |DIA: ** A **    VENC.: ** |"
    Mid(x_linha, 13, 36) = rsCliente("Razao Social").Value
    Mid(x_linha, 63, 14) = rsCliente("CPF").Value
    If Vencimento.LocalizarCodigo(rsCliente("Codigo do Vencimento").Value) Then
        Mid(x_linha, 116, 2) = Format(Vencimento.Dia2Inicial, "00")
        Mid(x_linha, 121, 2) = Format(Vencimento.Dia2Final, "00")
        Mid(x_linha, 134, 2) = Format(Vencimento.Dia2Vencimento, "00")
    End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|ENDERECO.:                                       |IDENT....:                     |                          |                          |"
    Mid(x_linha, 13, 36) = rsCliente("Endereco").Value
    Mid(x_linha, 63, 14) = rsCliente("Identidade").Value
    Mid(x_linha, 85, 10) = rsCliente("Orgao Emissor").Value
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|BAIRRO...:                                       |INSC.EST.:                     |                          |                          |"
    Mid(x_linha, 13, 30) = rsCliente("Bairro").Value
    If Val(rsCliente("Inscricao Estadual").Value) > 0 Then
        Mid(x_linha, 63, 14) = rsCliente("Inscricao Estadual").Value
    End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|CIDADE...:                               UF.:    |CGC......:                     |                          |                          |"
    Mid(x_linha, 13, 20) = rsCliente("Cidade").Value
    Mid(x_linha, 48, 2) = rsCliente("UF").Value
    If Val(rsCliente("CGC").Value) > 0 Then
        Mid(x_linha, 63, 18) = fMascaraCNPJ(rsCliente("CGC").Value)
    End If
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 6
End Sub
Private Sub ImpDetResumido()
    Dim x_linha As String
    Dim i As Integer
    
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    x_linha = "|         |                                          |                 |                     |                                          |"
    Mid(x_linha, 2, 6) = Format(rsCliente("Codigo").Value, "00,000")
    Mid(x_linha, 13, 40) = rsCliente("Razao Social").Value
    Mid(x_linha, 56, 14) = fMascaraTelefone(rsCliente("Telefone").Value)
    Mid(x_linha, 74, 14) = fMascaraCPF(rsCliente("CPF").Value)
    If Val(rsCliente("CGC").Value) > 0 Then
        Mid(x_linha, 74, 18) = fMascaraCNPJ(rsCliente("CGC").Value)
    End If
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim x_linha As String
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    If optRel1.Value = True Then
    Else
        If Me.optResumido.Value = True Then
            x_linha = "+---------+------------------------------------------+-----------------+---------------------+------------------------------------------+"
            Mid(x_linha, 14, 22) = " Cerrado Informática. "
        Else
            x_linha = "+-------------------------------------------------+-------------------------------+--------------------------+--------------------------+"
            Mid(x_linha, 4, 22) = " Cerrado Informática. "
        End If
    End If
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpCab()
    Dim x_linha As String
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
    x_linha = "|                                                                  Página,     |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    Mid(x_linha, 76, 3) = Format(lPagina, "000")
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| RELAÇÃO DE CLIENTES EM ORDEM ALFABÉTICA                  Goiânia,            |"
    Mid(x_linha, 69, 10) = msk_data
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    If optRel1.Value = True Then
    Else
        If Me.optResumido.Value = True Then
            BioImprime "@Printer.Print " & "+---------+------------------------------------------+-----------------+---------------------+------------------------------------------+"
            BioImprime "@Printer.Print " & "|  CODIGO | NOME DO CLIENTE                          |     TELEFONE    | CNPJ/CPF            |                                          |"
            BioImprime "@Printer.Print " & "+---------+------------------------------------------+-----------------+---------------------+------------------------------------------+"
        End If
    End If
End Sub
Private Sub cbo_vencimento_GotFocus()
    SendMessageLong cbo_vencimento.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_vencimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub chk_geral_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_vencimento.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = RetiraGString(1)
    chk_geral.SetFocus
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    lLocal = 1
    If ValidaCampos Then
        AtivaBotoes (False)
        If optRel1.Value = True Then
            If SelecionaImpressoraHP(Me) Then
                Call GravaAuditoria(1, Me.name, 6, "")
                RelatorioCR
            End If
        Else
            If SelecionaImpressoraEpson(Me) Then
                Call GravaAuditoria(1, Me.name, 7, "")
                Relatorio
            End If
        End If
        AtivaBotoes (True)
        cmd_sair.SetFocus
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf cbo_vencimento.ListIndex = -1 Then
        MsgBox "Selecione o vencimento", vbInformation, "Atenção!"
        cbo_vencimento.SetFocus
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
        AtivaBotoes (False)
        If optRel1.Value = True Then
            If SelecionaImpressoraHP(Me) Then
                Call GravaAuditoria(1, Me.name, 6, "")
                RelatorioCR
            End If
        Else
            If SelecionaImpressoraEpson(Me) Then
                Call GravaAuditoria(1, Me.name, 7, "")
                Relatorio
            End If
        End If
        AtivaBotoes (True)
        cmd_sair.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        chk_geral.Value = False
        cbo_vencimento.ListIndex = 0
        cmd_visualizar.SetFocus
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
    PreencheCboVencimento
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_geral.SetFocus
    End If
End Sub
