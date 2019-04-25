VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_auditoria 
   Caption         =   "Emissão de Auditoria"
   ClientHeight    =   4665
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_auditoria.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_auditoria.frx":030A
   ScaleHeight     =   4665
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_auditoria.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Visualiza inventário de produtos."
      Top             =   3660
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_auditoria.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Imprime inventário de produtos."
      Top             =   3660
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_auditoria.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3660
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   3435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CheckBox chkImprimeNomeComputador 
         Caption         =   "Imprimir Nome do Computador e Programa"
         Height          =   315
         Left            =   1680
         TabIndex        =   19
         Top             =   3000
         Width           =   4575
      End
      Begin VB.TextBox txtIP 
         Height          =   315
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   17
         Top             =   2340
         Width           =   1875
      End
      Begin VB.ComboBox cboPrograma 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1920
         Width           =   4755
      End
      Begin VB.ComboBox cboOperacaoAuditoria 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1500
         Width           =   4755
      End
      Begin VB.CheckBox chkImprimeDetalhe 
         Caption         =   "I&mprime detalhado"
         Height          =   315
         Left            =   1680
         TabIndex        =   18
         Top             =   2700
         Width           =   3015
      End
      Begin VB.ComboBox cboUsuario 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
         Width           =   4755
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_auditoria.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "lst_auditoria.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_auditoria.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
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
      Begin VB.Label Label3 
         Caption         =   "Número do IP"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   2340
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Nome do Programa"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Operação Auditoria"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Nome do Usuário"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "&Data de Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_auditoria"
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

Dim rs As New adodb.Recordset
Dim rstAuditoria As New adodb.Recordset

Private Cliente As New CadastroDLL.cCliente
Private MovimentoCupomFiscal As New CadastroDLL.cMovimentoCupomFiscal
Private Programa As New CadastroDLL.cPrograma

Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Cliente = Nothing
    Set MovimentoCupomFiscal = Nothing
    Set Programa = Nothing
    Set rstAuditoria = Nothing
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
End Sub
Private Sub PreencheCboOperacaoAuditoria()
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Codigo, Nome"
    lSQL = lSQL & "     FROM Operacao_Auditoria"
    lSQL = lSQL & " ORDER BY Nome, Codigo"
    'Abre RecordSet
    Set rs = Conectar.RsConexao(lSQL)
    
    cboOperacaoAuditoria.Clear
    cboOperacaoAuditoria.AddItem "Todas as Operações"
    cboOperacaoAuditoria.ItemData(cboOperacaoAuditoria.NewIndex) = 0
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do Until rs.EOF
            cboOperacaoAuditoria.AddItem rs("Nome").Value
            cboOperacaoAuditoria.ItemData(cboOperacaoAuditoria.NewIndex) = rs("Codigo").Value
            rs.MoveNext
        Loop
    End If
End Sub
Private Sub PreencheCboPrograma()
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Codigo, Tipo, [Nome para Menu] as Nome"
    lSQL = lSQL & "     FROM Programa"
    lSQL = lSQL & " ORDER BY Tipo, [Nome para Menu], Codigo"
    'Abre RecordSet
    Set rs = Conectar.RsConexao(lSQL)
    
    cboPrograma.Clear
    cboPrograma.AddItem "Todos os Programas"
    cboPrograma.ItemData(cboPrograma.NewIndex) = 0
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do Until rs.EOF
            cboPrograma.AddItem rs("Tipo").Value & " - " & rs("Nome").Value
            cboPrograma.ItemData(cboPrograma.NewIndex) = rs("Codigo").Value
            rs.MoveNext
        Loop
    End If
End Sub
Private Sub PreencheCboUsuario()
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Codigo, Nome"
    lSQL = lSQL & "     FROM Usuario"
    lSQL = lSQL & "    WHERE Situacao = " & preparaTexto("A")
    lSQL = lSQL & " ORDER BY Nome, Codigo"
    'Abre RecordSet
    Set rs = Conectar.RsConexao(lSQL)
    
    cboUsuario.Clear
    cboUsuario.AddItem "Todos os Usuários"
    cboUsuario.ItemData(cboUsuario.NewIndex) = 0
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do Until rs.EOF
            cboUsuario.AddItem rs("Nome").Value
            cboUsuario.ItemData(cboUsuario.NewIndex) = rs("Codigo").Value
            rs.MoveNext
        Loop
    End If
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    
    'Seleciona Produtos Vendidos dentro das condições
    lSQL = ""
    lSQL = lSQL & " SELECT Usuario.Nome As NomeUsuario, Auditoria.Data, Auditoria.Hora, "
    lSQL = lSQL & "        Programa.Tipo, Programa.[Nome para Menu], "
    lSQL = lSQL & "        Operacao_Auditoria.Nome As NomeAuditoria, Auditoria.Observacao, "
    lSQL = lSQL & "        Auditoria.Computador, Auditoria.[Nome Interno do Programa]"
    lSQL = lSQL & "   FROM Auditoria"
    lSQL = lSQL & "   LEFT JOIN Usuario ON Auditoria.[Codigo do Usuario] = Usuario.Codigo"
    lSQL = lSQL & "   LEFT JOIN Programa ON Auditoria.[Nome Interno do Programa] = Programa.[Nome Interno]"
    lSQL = lSQL & "   LEFT JOIN Operacao_Auditoria ON Auditoria.Operacao = Operacao_Auditoria.Codigo"
    lSQL = lSQL & "  WHERE Auditoria.Data >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "    AND Auditoria.Data <= " & preparaData(msk_data_f.Text)
    'lSQL = lSQL & "    AND Auditoria.[Codigo do Usuario] = Usuario.Codigo"
    'lSQL = lSQL & "    AND Auditoria.[Nome Interno do Programa] = Programa.[Nome Interno]"
    'lSQL = lSQL & "    AND Auditoria.Operacao = Operacao_Auditoria.Codigo"
    If cboUsuario.ListIndex > 0 Then
        lSQL = lSQL & "    AND [Codigo do Usuario] = " & Val(cboUsuario.ItemData(cboUsuario.ListIndex))
    End If
    If cboOperacaoAuditoria.ListIndex > 0 Then
        lSQL = lSQL & "    AND Operacao = " & Val(cboOperacaoAuditoria.ItemData(cboOperacaoAuditoria.ListIndex))
    End If
    If cboPrograma.ListIndex > 0 Then
        If Programa.LocalizarNomeMenu(Mid(cboPrograma.Text, 1, 2), Mid(cboPrograma.Text, 6, Len(cboPrograma.Text) - 5)) Then
            lSQL = lSQL & "    AND [Nome Interno do Programa] = " & preparaTexto(Programa.NomeInterno)
        Else
            MsgBox "Não foi possível localizar o programa: " & cboPrograma.Text, vbCritical, "Erro de Integridade!"
        End If
    End If
    
    'Teste para relatorio pro Ivander
    If cboPrograma.Text = "MO - Baixa Nota de Abastecimento P/ Período" Or cboPrograma.Text = "MO - Baixa Nota de Abastecimento (Individual)" Then
        If cboOperacaoAuditoria.ListIndex = 0 Then
            '18 Baixar, 19-Estornar, 10-Confirmar
            lSQL = lSQL & "    AND ( Operacao = 18 OR Operacao = 19 OR Operacao = 10)"
        End If
    End If
    
'    If cboTipoVenda.ListIndex = 0 Then
'        lSQl = lSQl & "    AND Produto.[Exclusivo Posto] = " & preparaBooleano(True)
'    End If
'    If cboTipoVenda.ListIndex = 1 Then
'        lSQl = lSQl & "    AND Produto.[Exclusivo Loja] = " & preparaBooleano(True)
'    End If
    If txtIP.Text <> "" Then
        lSQL = lSQL & "    AND Auditoria.Computador LIKE " & preparaTexto("%" & txtIP.Text & "%")
    End If
    lSQL = lSQL & "  ORDER BY Auditoria.Data, Auditoria.Hora"
    Set rstAuditoria = Conectar.RsConexao(lSQL)
    If rstAuditoria.RecordCount > 0 Then
        ImpDados
    End If
    rstAuditoria.Close
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    LoopAuditoria
    If lPagina > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Auditoria|@|"
        'gStringChamada = g_string
        'menu_personalizado.GravaSgpNetCadastroIni ("preview")
        'gStringChamada = ""
        frm_preview.Show 1
    End If
End Sub
Private Sub LoopAuditoria()
    If rstAuditoria.RecordCount > 0 Then
        Do Until rstAuditoria.EOF
            ImpDet
            rstAuditoria.MoveNext
        Loop
    End If
End Sub
Private Sub ImpDet()
    Dim xLinha As String
    Dim i As Integer
    
    If lPagina = 0 Then
        ImpCab
    End If
    If lLinha >= 60 Then
        xLinha = "+----------+--------+--------------------------------+--------------------------------+----+--------------------------------------------+"
        Mid(xLinha, 16, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    xLinha = "|          |        |                                |                                |    |                                            |"

    '                   1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '          12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    '          |99/99/9999|00:00:00| 123456789012345678901234567890 | 123456789012345678901234567890 | 12 | 1234567890123456789012345678901234567890   |
    '          | 123456789012345678901234567890                     | 12345678901234567890123456789012345678901234567890123456789012345678901234567890 |
    
    Mid(xLinha, 2, 10) = Format(rstAuditoria!Data, "dd/MM/yyyy")
    Mid(xLinha, 13, 8) = Format(rstAuditoria!Hora, "hh:mm:ss")
    If IsNull(rstAuditoria!NomeUsuario) Then
        Mid(xLinha, 23, 30) = "** Não Logado **"
    Else
        Mid(xLinha, 23, 30) = rstAuditoria!NomeUsuario
    End If
    If chkImprimeDetalhe.Value = 1 Then
        Mid(xLinha, 56, 30) = rstAuditoria!Computador
    Else
        Mid(xLinha, 56, 30) = "" & rstAuditoria!NomeAuditoria
    End If
    If IsNull(rstAuditoria!Tipo) Then
        Mid(xLinha, 89, 2) = "**"
    Else
        Mid(xLinha, 89, 2) = rstAuditoria!Tipo
    End If
    If IsNull(rstAuditoria![Nome para Menu]) Then
        Mid(xLinha, 94, 40) = "** " & rstAuditoria![Nome Interno do Programa] & " **"
    Else
        Mid(xLinha, 94, 40) = rstAuditoria![Nome para Menu]
    End If
    
    If chkImprimeNomeComputador.Value = 0 Then
        Mid(xLinha, 54, 84) = "|                                                                                  |"
    End If
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
    
    If chkImprimeDetalhe.Value = 1 Then
        xLinha = "|                                                    |                                                                                  |"
        Mid(xLinha, 3, 30) = "" & rstAuditoria!NomeAuditoria
        Mid(xLinha, 56, 80) = rstAuditoria!Observacao
        'Teste para relatorio pro Ivander
        If cboPrograma.Text = "MO - Baixa Nota de Abastecimento P/ Período" Or cboPrograma.Text = "MO - Baixa Nota de Abastecimento (Individual)" Then
            If cboOperacaoAuditoria.ListIndex = 0 And Mid(rstAuditoria!Observacao, 1, 4) = "Cli:" Then
                For i = 5 To 10
                    If Mid(rstAuditoria!Observacao, i, 1) = " " Then
                        Exit For
                    End If
                Next
                If Cliente.LocalizarCodigo(Val(Mid(rstAuditoria!Observacao, 5, i - 5))) Then
                    Mid(xLinha, 23, 30) = Cliente.RazaoSocial
                End If
            End If
        End If
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
        xLinha = "+----------+--------+--------------------------------+--------------------------------+----+--------------------------------------------+"
        If chkImprimeNomeComputador.Value = 1 Then
            BioImprime "@Printer.Print " & xLinha
            lLinha = lLinha + 1
        End If
    End If
End Sub
Private Sub ImpTotal()
'    Dim y_local As Single
    Dim xLinha As String
    Dim i As Integer
    If chkImprimeDetalhe.Value = 0 Then
        xLinha = "+----------+--------+--------------------------------+--------------------------------+----+--------------------------------------------+"
        Mid(xLinha, 16, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.FontName = Courier New"
        BioImprime "@Printer.Print " & " "
    Else
        BioImprime "@@Printer.FontName = Courier New"
        BioImprime "@Printer.Print " & " "
    End If
End Sub
Private Sub ImpCab()
    Dim xLinha As String
    Dim i As Integer
    If lPagina = 0 Then
        lNomeArquivo = BioCriaImprime
        'seleciona medidas para centímetros
        BioImprime "@@Printer.ScaleMode = 7"
        BioImprime "@@Printer.PaperSize = 1"
        BioImprime "@@Printer.FontName = Courier New"
        BioImprime "@@Printer.FontName = Courier New"
        'teste para imprimir letra correta
        BioImprime "@@Printer.FontBold = False"
        BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    End If
    lPagina = lPagina + 1
    lLinha = 0
    BioImprime "@@Printer.FontName = Draft 5cpi"
    BioImprime "@@Printer.FontName = Courier New"
    BioImprime "@@Printer.CurrentY = 0"
    BioImprime "@@Printer.Print " & Chr(34) & " " & Chr(34)
    BioImprime "@@Printer.FontName = Courier New"
    BioImprime "@@Printer.FontSize = 10"
    BioImprime "@@Printer.CurrentY = 0"
    '                  1         2         3         4         5         6         7         8
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "+------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = True"
    xLinha = "|                                                                  Página: ___ |"
    Mid(xLinha, 3, 40) = g_nome_empresa
    Mid(xLinha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    xLinha = "| AUDITORIA DO SISTEMA                                            , __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| REFERENTE AO PERIODO DE.: __/__/____ A __/__/____.                           |"
    Mid(xLinha, 29, 10) = msk_data_i.Text
    Mid(xLinha, 42, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & xLinha
'    xLinha = "| CAIXA INICIAL...........: X    CAIXA FINAL..: X                              |"
'    Mid(xLinha, 29, 1) = cbo_periodo_i.Text
'    Mid(xLinha, 49, 1) = cbo_periodo_f.Text
'    BioImprime "@Printer.Print " & xLinha
    xLinha = "| USUARIO.................:                                                    |"
    Mid(xLinha, 29, 40) = cboUsuario.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| OPERACAO DE AUDITORIA...:                                                    |"
    Mid(xLinha, 29, 30) = cboOperacaoAuditoria.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| NOME DO PROGRAMA........:                                                    |"
    Mid(xLinha, 29, 40) = cboPrograma.Text
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Courier New"
    BioImprime "@@Printer.FontSize = 7"
    xLinha = "+----------+--------+--------------------------------+--------------------------------+----+--------------------------------------------+"
    BioImprime "@Printer.Print " & xLinha
    If chkImprimeDetalhe.Value = 0 Then
        xLinha = "|   DATA   |  HORA  | NOME DO USUARIO                | OPERACAO                       | TP | PROGRAMA                                   |"
        BioImprime "@Printer.Print " & xLinha
    End If
    If chkImprimeDetalhe.Value = 1 Then
        xLinha = "|   DATA   |  HORA  | NOME DO USUARIO                | NOME DO COMPUTADOR             | TP | PROGRAMA                                   |"
        BioImprime "@Printer.Print " & xLinha
        xLinha = "| OPERACAO                                           | OBSERVACAO                                                                       |"
        BioImprime "@Printer.Print " & xLinha
    End If
    xLinha = "+----------+--------+--------------------------------+--------------------------------+----+--------------------------------------------+"
    BioImprime "@Printer.Print " & xLinha
End Sub
Private Sub cboOperacaoAuditoria_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboPrograma.SetFocus
    End If
End Sub
Private Sub cboPrograma_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub cboUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboOperacaoAuditoria.SetFocus
    End If
End Sub
Private Sub cboTipoVenda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboUsuario.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
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
    g_string = msk_data_f.Text
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
    g_string = msk_data_i.Text
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
        If SelecionaImpressoraHP(Me) Then
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
    ElseIf cboOperacaoAuditoria.ListIndex = -1 Then
        MsgBox "Selecione uma forma de pagamento.", vbInformation, "Atenção!"
        cboOperacaoAuditoria.SetFocus
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
        If SelecionaImpressoraHP(Me) Then
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
        msk_data_f.Text = Format(g_data_def, "dd/mm/yyyy")
        cboUsuario.ListIndex = 0
        cboOperacaoAuditoria.ListIndex = 0
        cboPrograma.ListIndex = 0
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
    PreencheCboUsuario
    PreencheCboOperacaoAuditoria
    PreencheCboPrograma
    chkImprimeNomeComputador.Value = 1
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
        cboUsuario.SetFocus
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
