VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form lst_entrada_produto_conferencia 
   Caption         =   "Emissão das Entradas de Produtos (Conferência)"
   ClientHeight    =   2295
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6210
   Icon            =   "lst_entrada_produto_conferencia.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_entrada_produto_conferencia.frx":030A
   ScaleHeight     =   2295
   ScaleWidth      =   6210
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1020
      Picture         =   "lst_entrada_produto_conferencia.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Visualiza entrada de produtos para conferência."
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2700
      Picture         =   "lst_entrada_produto_conferencia.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprime entrada de produtos para conferência."
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4380
      Picture         =   "lst_entrada_produto_conferencia.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1320
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1155
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5955
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5340
         Picture         =   "lst_entrada_produto_conferencia.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_entrada_produto_conferencia.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_entrada_produto_conferencia.frx":6CBA
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
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4320
         TabIndex        =   8
         Top             =   660
         Width           =   975
         _ExtentX        =   1720
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
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   315
         Left            =   3540
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1515
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
      Left            =   120
      Top             =   1860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_entrada_produto_conferencia"
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
Dim lQuantidade(1 To 7) As Currency
Dim lTotal(1 To 7) As Currency
Dim lSQL As String

Private Produto As New cProduto
Private rsTabela As New adodb.Recordset
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Produto = Nothing
End Sub
Private Sub ZeraVariaveis()
    Dim i As Integer
    
    lLinha = 0
    lPagina = 0
    For i = 1 To 7
        lQuantidade(i) = 0
        lTotal(i) = 0
    Next
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    Call LoopEmpresa
    If lPagina > 0 Then
        Call ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Entrada de Produto p/ Conferência|@|"
        frm_preview.Show 1
    End If
    cmd_sair.SetFocus
End Sub
Private Sub LoopEmpresa()
    Dim i As Integer
    Dim rstEmpresa As New adodb.Recordset
    
    Set rstEmpresa = Conectar.RsConexao("SELECT Codigo, Nome FROM Empresas WHERE Inativo = " & preparaBooleano(False) & "ORDER BY Codigo")
    With rstEmpresa
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                i = !Codigo
                ImpMovimentoEntradaProduto i
                .MoveNext
            Loop
        End If
    End With
    rstEmpresa.Close
    Set rstEmpresa = Nothing
End Sub
Private Sub ImpMovimentoEntradaProduto(ByVal pEmpresa As Integer)
    lSQL = ""
    lSQL = lSQL & "SELECT [Data da Entrada], [Numero do Documento], [Codigo do Produto], Quantidade, [Preco de Custo], Observacao"
    lSQL = lSQL & "  FROM Entrada_Produto"
    lSQL = lSQL & " WHERE Empresa = " & pEmpresa
    lSQL = lSQL & "   AND [Data da Digitacao] >= " & preparaData(msk_data_i.Text)
    lSQL = lSQL & "   AND [Data da Digitacao] <= " & preparaData(msk_data_f.Text)
    lSQL = lSQL & "   AND [Tipo da Entrada] <> " & 3
    lSQL = lSQL & " ORDER BY [Data da Digitacao], [Codigo do Produto]"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            If Produto.LocalizarCodigo(rsTabela("Codigo do Produto").Value) Then
                Call ImpDet(rsTabela("Data da Entrada").Value, rsTabela("Numero do Documento").Value, rsTabela("Codigo do Produto").Value, Produto.Nome, Produto.Unidade, rsTabela("Quantidade").Value, rsTabela("Preco de Custo").Value, Format(rsTabela("Quantidade").Value * rsTabela("Preco de Custo").Value, "00000000.00"), rsTabela("Observacao").Value, pEmpresa)
                lTotal(pEmpresa) = lTotal(pEmpresa) + Format(rsTabela("Quantidade").Value * rsTabela("Preco de Custo").Value, "00000000.00")
                lTotal(7) = lTotal(7) + Format(rsTabela("Quantidade").Value * rsTabela("Preco de Custo").Value, "00000000.00")
                lQuantidade(pEmpresa) = lQuantidade(pEmpresa) + rsTabela("Quantidade").Value
                lQuantidade(7) = lQuantidade(7) + rsTabela("Quantidade").Value
            Else
                MsgBox "Produto inexistente!" & Chr(10) & rsTabela("Codigo do Produto").Value, vbInformation, "Erro de integridade!"
            End If
            rsTabela.MoveNext
        Loop
        ImpSubTotal pEmpresa
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
'    With tbl_entrada_produto
'        .Seek ">=", i, CDate(msk_data_i.Text), CDate("01/01/1900"), 0, 0
'        If Not .NoMatch Then
'            Do Until .EOF
'                If !Empresa <> i Or ![Data da Digitacao] > CDate(msk_data_f.Text) Then
'                    Exit Do
'                End If
'                If ![Tipo da Entrada] <> 3 Then
'                    If Produto.LocalizarCodigo(![Codigo do Produto]) Then
'                        Call ImpDet(![Data da Entrada], ![Numero do Documento], ![Codigo do Produto], Produto.Nome, Produto.unidade, !Quantidade, ![Preco de Custo], Format(!Quantidade * ![Preco de Custo], "00000000.00"), !Observacao, i)
'                        lTotal(i) = lTotal(i) + Format(!Quantidade * ![Preco de Custo], "00000000.00")
'                        lTotal(7) = lTotal(7) + Format(!Quantidade * ![Preco de Custo], "00000000.00")
'                        lQuantidade(i) = lQuantidade(i) + !Quantidade
'                        lQuantidade(7) = lQuantidade(7) + !Quantidade
'                    Else
'                        MsgBox "Produto inexistente!" & Chr(10) & ![Codigo do Produto], vbInformation, "Erro de integridade!"
'                    End If
'                End If
'                .MoveNext
'            Loop
'        End If
'    End With
End Sub
Private Sub ImpDet(ByVal pData As Date, ByVal pDocumento As String, ByVal pCodigo As Long, ByVal pNome As String, ByVal pUnidade As String, ByVal pQuantidade As Currency, ByVal pValorUnitario As Currency, ByVal pValorTotal As Currency, ByVal pObservacao As String, ByVal pEmpresa As Integer)
    Dim xLinha As String
    Dim i As Integer
    
    If lPagina = 0 Then
        ImpCab
    End If
    If lLinha >= 60 Then
        xLinha = "+----------+----------+------+---------------------------------------+---+---------+-------------+-------------+------------------------+"
        Mid(xLinha, 34, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    xLinha = "|          |          |      |                                       |   |         |             |             |                     |  |"
    Mid(xLinha, 2, 10) = Format(pData, "dd/mm/yyyy")
    Mid(xLinha, 13, 10) = pDocumento
    i = Len(Format(pCodigo, "##000"))
    Mid(xLinha, 24 + 5 - i, i) = Format(pCodigo, "##000")
    Mid(xLinha, 31, 39) = pNome
    Mid(xLinha, 71, 3) = pUnidade
    i = Len(Format(pQuantidade, "##,##0.00"))
    Mid(xLinha, 75 + 9 - i, i) = Format(pQuantidade, "##,##0.00")
    i = Len(Format(pValorUnitario, "#####,##0.00"))
    Mid(xLinha, 85 + 12 - i, i) = Format(pValorUnitario, "#####,##0.00")
    i = Len(Format(pValorTotal, "#####,##0.00"))
    Mid(xLinha, 99 + 12 - i, i) = Format(pValorTotal, "#####,##0.00")
    Mid(xLinha, 113, 21) = pObservacao
    i = Len(Format(pEmpresa, "#0"))
    Mid(xLinha, 135 + 2 - i, i) = Format(pEmpresa, "#0")
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpSubTotal(ByVal pEmpresa As Integer)
    Dim y_local As Single
    Dim xLinha As String
    Dim i As Integer
    
    xLinha = "|          |          |      |                   ** Total da Empresa |   |         |             |             |                     |  |"
    i = Len(Format(lQuantidade(pEmpresa), "##,##0.00"))
    Mid(xLinha, 75 + 9 - i, i) = Format(lQuantidade(pEmpresa), "##,##0.00")
    i = Len(Format(lTotal(pEmpresa), "#####,##0.00"))
    Mid(xLinha, 99 + 12 - i, i) = Format(lTotal(pEmpresa), "#####,##0.00")
    i = Len(Format(pEmpresa, "#0"))
    Mid(xLinha, 135 + 2 - i, i) = Format(pEmpresa, "#0")
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@y_local = Printer.CurrentY"
    BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    BioImprime "@@Printer.CurrentY = y_local"
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.CurrentY = y_local"
    BioImprime "@@Printer.Print " & "  "
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+----------+----------+------+---------------------------------------+---+---------+-------------+-------------+---------------------+--+"
End Sub
Private Sub ImpTotal()
    Dim y_local As Single
    Dim xLinha As String
    Dim i As Integer
    
    xLinha = "|                                                *** Total Geral         |         |             |             |                        |"
    i = Len(Format(lQuantidade(7), "##,##0.00"))
    Mid(xLinha, 75 + 9 - i, i) = Format(lQuantidade(7), "##,##0.00")
    i = Len(Format(lTotal(7), "#####,##0.00"))
    Mid(xLinha, 99 + 12 - i, i) = Format(lTotal(7), "#####,##0.00")
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@y_local = Printer.CurrentY"
    BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    BioImprime "@@Printer.CurrentY = y_local"
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.CurrentY = y_local"
    BioImprime "@@Printer.Print " & "  "
    BioImprime "@@Printer.FontBold = False"
    xLinha = "+------------------------------------------------------------------------+---------+-------------+-------------+------------------------+"
    Mid(xLinha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & "  "
End Sub
Private Sub ImpCab()
    Dim xLinha As String
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
    BioImprime "@@Printer.Print " & "  "
    Printer.FontName = "Sans Serif 10cpi"
    BioImprime "@@Printer.CurrentY = 0"
    BioImprime "@Printer.Print " & "+------------------------------------------------------------------------------+"
    BioImprime "@@Printer.FontBold = True"
    xLinha = "| GRUPO X                                                          Página: ___ |"
    g_string = ReadINI("GRUPO DE EMPRESAS", "Nome do Grupo", gArquivoIni)
    Mid(xLinha, 3, 40) = g_string
    g_string = ""
    Mid(xLinha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    xLinha = "| ENTRADAS DE PRODUTOS (CONFERÊNCIA)                              , __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| PERÍODO DA DIGITAÇÃO.: __/__/____ A __/__/____                               |"
    Mid(xLinha, 26, 10) = msk_data_i.Text
    Mid(xLinha, 39, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & "+----------+----------+------+---------------------------------------+---+---------+-------------+-------------+---------------------+--+"
    BioImprime "@Printer.Print " & "|  D A T A |DOCUMENTO |CODIGO|DISCRIMINACAO DOS PRODUTOS             |UN.|  QUANT. |VLR. UNITARIO| TOTAL CUSTO |OBSERVACAO           |EM|"
    BioImprime "@Printer.Print " & "+----------+----------+------+---------------------------------------+---+---------+-------------+-------------+---------------------+--+"
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
    g_string = ""
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
    g_string = ""
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
    g_string = ""
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
        MsgBox "Informe a data de emissão.", vbInformation, "Dados Incompleto!"
        msk_data.SetFocus
    ElseIf Not IsDate(msk_data_i.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Dados Incompleto!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Dados Incompleto!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f.Text) < CDate(msk_data_i.Text) Then
        MsgBox "Data final deve ser maior que a data inicial.", vbInformation, "Dados Incompleto!"
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
        msk_data_i.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def, "dd/mm/yyyy")
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
        cmd_visualizar.SetFocus
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
