VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_cliente_cheque 
   Caption         =   "Emissão dos Clientes (Cheque)"
   ClientHeight    =   1905
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   7290
   Icon            =   "emissao_cliente_cheque.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "emissao_cliente_cheque.frx":030A
   ScaleHeight     =   1905
   ScaleWidth      =   7290
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1320
      Picture         =   "emissao_cliente_cheque.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Visualiza os clientes em ordem alfabética."
      Top             =   960
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3240
      Picture         =   "emissao_cliente_cheque.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprime os clientes em ordem alfabética."
      Top             =   960
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5160
      Picture         =   "emissao_cliente_cheque.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   960
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7035
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   3060
         Picture         =   "emissao_cliente_cheque.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
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
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_cliente_cheque"
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

Private lConnCheq As adodb.Connection
Private rsCliente As New adodb.Recordset
Private rsAuxiliar As New adodb.Recordset
Private Sub Finaliza()
    lConnCheq.Close
    Set lConnCheq = Nothing
    Call GravaAuditoria(1, Me.name, 11, "")
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    
    'Prepara SQL
    lSQL = "SELECT Codigo, Nome, [Limite de Credito], Prazo, Aprovado, Condicao,"
    lSQL = lSQL & " Inativo, Tipo"
    lSQL = lSQL & " FROM Pessoa"
    lSQL = lSQL & " WHERE Inativo = " & preparaBooleano(False)
    lSQL = lSQL & " ORDER BY Nome"
    
    'Abre RecordSet
    Set rsCliente = New adodb.Recordset
    Set rsCliente = lConnCheq.Execute(lSQL)
    
    
    'Verifica movimento
    If rsCliente.EOF = False Then
        ImpDados
    End If
    If rsCliente.State = 1 Then
        rsCliente.Close
    End If
    BioImprime "@@Printer.EndDoc"
    BioFechaImprime
    g_string = lLocal & lNomeArquivo & "|@|Relatório de Clientes (Cheque)|@|"
    frm_preview.Show 1
    
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    Dim x_linha As String
    
    'loop cliente
    Do Until rsCliente.EOF
        If lPagina = 0 Then
            ImpCab
        End If
        If lLinha >= 55 Then
            x_linha = "+---------+------------------------------------------+------------+-----+-----+-----+-----------------+---------------------------------+"
            Mid(x_linha, 15, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & x_linha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        If rsCliente("Tipo").Value = 1 Then
            lSQL = "SELECT CPF"
            lSQL = lSQL & " FROM PessoaFisica"
            lSQL = lSQL & " WHERE Codigo = " & rsCliente("Codigo").Value
        Else
            lSQL = "SELECT CNPJ"
            lSQL = lSQL & " FROM PessoaJuridica"
            lSQL = lSQL & " WHERE Codigo = " & rsCliente("Codigo").Value
        End If
        Set rsAuxiliar = New adodb.Recordset
        Set rsAuxiliar = lConnCheq.Execute(lSQL)
        ImpDet
        rsAuxiliar.Close
        rsCliente.MoveNext
    Loop
    ImpTotal
End Sub
Private Sub ImpDet()
    Dim x_linha As String
    Dim i As Integer
    '                   1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '          12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    '          |  CODIGO | NOME DO CLIENTE                          |   CREDITO  |APROV|INAT.|PRAZO|     TELEFONE    |                                 |"
    x_linha = "| 000.000 | 1234567890123456789012345678901234567890 | 000.000,00 | sim | sit |  00 | (   ) 5555-5555 |                                 |"
    x_linha = "|         |                                          |            |     |     |     |                 |                                 |"
    Mid(x_linha, 3, 7) = Format(rsCliente("Codigo").Value, "000,000")
    Mid(x_linha, 13, 40) = rsCliente("Nome").Value
    i = Len(Format(rsCliente("Limite de Credito").Value, "###,##0.00"))
    Mid(x_linha, 56 + 10 - i, i) = Format(rsCliente("Limite de Credito").Value, "###,##0.00")
    If UCase(rsCliente("Aprovado").Value) = "S" Then
        Mid(x_linha, 69, 3) = "SIM"
    Else
        Mid(x_linha, 69, 3) = "NAO"
    End If
    If rsCliente("Inativo").Value Then
        Mid(x_linha, 75, 3) = "SIM"
    Else
        Mid(x_linha, 75, 3) = "NAO"
    End If
    Mid(x_linha, 82, 2) = Format(rsCliente("Prazo").Value, "00")
    If rsAuxiliar.EOF = False Then
        If rsCliente("Tipo").Value = 1 Then
            Mid(x_linha, 105, 20) = fMascaraCPF(rsAuxiliar("CPF").Value)
        Else
            Mid(x_linha, 105, 20) = fMascaraCNPJ(rsAuxiliar("CNPJ").Value)
        End If
    End If
    BioImprime "@Printer.Print " & x_linha
    
    
'    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
'    x_linha = "+-------------------------------------------------+-------------------------------+--------------------------+--------------------------+"
'    BioImprime "@Printer.Print " & x_linha
'    x_linha = "|CODIGO...:                                       |C.E.P....:       -             |FONE.: (   )    -         |DIA: ** A **    VENC.: ** |"
'    Mid(x_linha, 13, 6) = Format(rsCliente("Codigo").Value, "00,000")
    'Mid(x_linha, 63, 6) = Format(Mid(rsCliente("CEP").Value, 1, 5), "00,000")
    'Mid(x_linha, 70, 4) = Mid(rsCliente("CEP").Value, 6, 3)
    'Mid(x_linha, 92, 11) = fMascaraTelefone(rsCliente("Telefone").Value)
'    BioImprime "@Printer.Print " & x_linha
'    x_linha = "|NOME.....:                                       |CPF......:                     |                          |DIA: ** A **    VENC.: ** |"
'    Mid(x_linha, 13, 36) = rsCliente("Nome").Value
    'Mid(x_linha, 63, 14) = rsCliente("CPF").Value
'    BioImprime "@Printer.Print " & x_linha
'    x_linha = "|ENDERECO.:                                       |IDENT....:                     |                          |                          |"
    'Mid(x_linha, 13, 36) = rsCliente("Endereco").Value
    'Mid(x_linha, 63, 14) = rsCliente("Identidade").Value
    'Mid(x_linha, 85, 10) = rsCliente("Orgao Emissor").Value
'    BioImprime "@Printer.Print " & x_linha
'    x_linha = "|BAIRRO...:                                       |INSC.EST.:                     |                          |                          |"
    'Mid(x_linha, 13, 30) = rsCliente("Bairro").Value
    'If Val(rsCliente("Inscricao Estadual").Value) > 0 Then
    '    Mid(x_linha, 63, 14) = rsCliente("Inscricao Estadual").Value
    'End If
'    BioImprime "@Printer.Print " & x_linha
'    x_linha = "|CIDADE...:                               UF.:    |CGC......:                     |                          |                          |"
    'Mid(x_linha, 13, 20) = rsCliente("Cidade").Value
    'Mid(x_linha, 48, 2) = rsCliente("UF").Value
    'If Val(rsCliente("CGC").Value) > 0 Then
    '    Mid(x_linha, 63, 18) = fMascaraCNPJ(rsCliente("CGC").Value)
    'End If
'    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim x_linha As String
    
    x_linha = "+---------+------------------------------------------+------------+-----+-----+-----+-----------------+---------------------------------+"
    Mid(x_linha, 15, 22) = " Cerrado Informática. "
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
    Mid(x_linha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+---------+------------------------------------------+------------+-----+-----+-----+-----------------+---------------------------------+"
    BioImprime "@Printer.Print " & "|  CODIGO | NOME DO CLIENTE                          |   CREDITO  |APROV|INAT.|PRAZO|     TELEFONE    | CNPJ/CPF                        |"
    BioImprime "@Printer.Print " & "+---------+------------------------------------------+------------+-----+-----+-----+-----------------+---------------------------------+"
End Sub
Private Sub chk_geral_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = RetiraGString(1)
    cmd_visualizar.SetFocus
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
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
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
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
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
    Dim xNomeArquivo As String
    Dim xSGBD As String
    Dim xNomeBancoDados As String
    Dim xDrive As String
    Dim xDiretorio As String
    Dim xStringConexao As String
    
    CentraForm Me
    xNomeArquivo = "c:\CheqPosto.ini"
    xSGBD = ReadINI("SGBD", "Gerenciador de Banco de Dados", xNomeArquivo)
    xDrive = ReadINI("Local", "Drive", xNomeArquivo)
    xDiretorio = ReadINI("Local", "Diretorio BD", xNomeArquivo)
    xNomeBancoDados = ReadINI("Local", "Nome do Banco de Dados", xNomeArquivo)
    
    If xSGBD = "ACCESS" Then
        xStringConexao = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & xDrive & xDiretorio & xNomeBancoDados
    ElseIf xSGBD = "SQLSERVER" Then
        xStringConexao = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & xNomeBancoDados & ";INITIAL CATALOG=" & "CheqPosto_Data" & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
    End If
    Set lConnCheq = New adodb.Connection
    lConnCheq.ConnectionString = xStringConexao
    
    lConnCheq.Open
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
