VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_grps 
   Caption         =   "Emissão de GPS"
   ClientHeight    =   2175
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   5535
   Icon            =   "Lst_grps.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Lst_grps.frx":030A
   ScaleHeight     =   2175
   ScaleWidth      =   5535
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   1380
      Picture         =   "Lst_grps.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprime a guia de GPS."
      Top             =   1200
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3360
      Picture         =   "Lst_grps.frx":1D5A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1200
      Width           =   795
   End
   Begin VB.Frame frmDados
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5295
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   3000
         Picture         =   "Lst_grps.frx":33EC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_valor 
         Height          =   300
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   300
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   7
         Format          =   "mm/yyyy"
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "&Empregadores/Autônomos"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "&Mês/Ano"
         Height          =   300
         Index           =   3
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_grps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_ano_mes As String * 6
Dim l_codigo As Integer
Dim l_funcionario As Integer
Dim l_local As Integer
Dim lColunaI As Currency
Dim lLinhaI As Currency
Dim lTotalEmpregado As Currency
Dim lQtdEmpregado As Integer
Dim lValorAutonomo As Currency
Dim lTotalSegurado As Currency
Dim lSalarioFamilia As Currency
Dim tbl_empresa As Table
Dim tbl_funcionario As Table
Dim tbl_movimento_falta_funcionario As Table
Dim tbl_movimento_folha As Table
Dim tbl_tabela_folha As Table
Dim tbl_tabela_provento_desconto As Table
Function CalculaValorEmpregado() As Currency
    CalculaValorEmpregado = 0
    lTotalEmpregado = 0
    lQtdEmpregado = 0
    With tbl_movimento_folha
        If .RecordCount > 0 Then
            .Seek ">=", g_empresa, l_ano_mes, 0, 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Or ![Mes Ano] <> l_ano_mes Then
                        Exit Do
                    End If
                    tbl_funcionario.Seek "=", g_empresa, ![Codigo do Funcionario]
                    If Not tbl_funcionario.NoMatch Then
                        If tbl_funcionario![Serie da Carteira de Trabalho] <> "NR" Then
                            If ![Codigo do Movimento] <> 15 Then
                                If ![Codigo do Movimento] < 500 Then
                                    CalculaValorEmpregado = CalculaValorEmpregado + !valor
                                    lTotalEmpregado = lTotalEmpregado + !valor
                                End If
                            Else
                                lSalarioFamilia = lSalarioFamilia + !valor
                            End If
                            If ![Codigo do Movimento] = 505 Or ![Codigo do Movimento] = 510 Or ![Codigo do Movimento] = 515 Then
                                CalculaValorEmpregado = CalculaValorEmpregado - !valor
                                lTotalEmpregado = lTotalEmpregado - !valor
                            End If
                            If ![Codigo do Movimento] = 1 Then
                                lQtdEmpregado = lQtdEmpregado + 1
                            ElseIf ![Codigo do Movimento] = 520 Then
                                lTotalSegurado = lTotalSegurado + !valor
                            End If
                        End If
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
End Function
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_empresa.Close
    tbl_funcionario.Close
    tbl_movimento_falta_funcionario.Close
    tbl_movimento_folha.Close
    tbl_tabela_folha.Close
    tbl_tabela_provento_desconto.Close
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = Mid(RetiraGString(1), 4, 7)
    txt_valor.SetFocus
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    If SelecionaImpressoraHP(Me) Then
        Call GravaAuditoria(1, Me.name, 7, "")
        Relatorio
    End If
End Sub
Private Sub ImpGPS()
    Dim x_cgc As String
    Dim x_empresa As Currency
    Dim x_terceiro As Currency
    Dim x_total As Currency
    l_local = 1
    lColunaI = -0.7
    lLinhaI = 0
    'seleciona medidas para polegadas
    Printer.ScaleMode = 5
    'Seleciona largura do formulário
    Printer.ScaleWidth = 8
'    largura_form = Printer.ScaleWidth
    'Seleciona altura do formulário
    Printer.ScaleHeight = 5.5
'    tamanho_form = Printer.ScaleHeight
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    'Seleciona nome da fonte
    Printer.FontName = "Arial"
    Printer.FontName = "Roman 10cpi"
    tbl_empresa.Seek "=", g_empresa
    If Not tbl_empresa.NoMatch Then
        x_cgc = Mid(tbl_empresa!CGC, 1, 2) & "."
        x_cgc = x_cgc & Mid(tbl_empresa!CGC, 3, 3) & "."
        x_cgc = x_cgc & Mid(tbl_empresa!CGC, 6, 3) & "/"
        x_cgc = x_cgc & Mid(tbl_empresa!CGC, 9, 4) & "-"
        x_cgc = x_cgc & Mid(tbl_empresa!CGC, 13, 2)
    End If
    
    Printer.FontName = "Arial"
    Printer.FontSize = 12
    Printer.FontBold = True
'    ImprimeCentralizado x_cgc, lColunaI + 3, lColunaI + 9.8, lLinhaI + 0.7, l_local
    ImprimeCentralizado x_cgc, lColunaI + 15.1, lColunaI + 19.4, lLinhaI + 2, l_local
'    Printer.FontName = "Arial"
'    Printer.FontSize = 12
'    Printer.FontBold = True
'    ImprimeCentralizado UCase(tbl_empresa!Nome), lColunaI + 3, lColunaI + 9.8, lLinhaI + 1.5, l_local
'    Printer.FontName = "Arial"
'    Printer.FontSize = 10
'    Printer.FontBold = True
'    ImprimeCentralizado Trim(tbl_empresa!Endereco) & " - " & Trim(tbl_empresa!Bairro), lColunaI + 3, lColunaI + 9.8, lLinhaI + 2.2, l_local
'    ImprimeCentralizado "CEP  -  " & Mid(tbl_empresa!CEP, 1, 2) & "." & Mid(tbl_empresa!CEP, 3, 3) & "-" & Mid(tbl_empresa!CEP, 6, 3), lColunaI + 3, lColunaI + 9.8, lLinhaI + 2.9, l_local
'    Printer.FontSize = 12
'    Printer.FontBold = True
'    ImprimeCentralizado Trim(tbl_empresa!Cidade) & "  -  " & Trim(tbl_empresa!estado), lColunaI + 3, lColunaI + 9.8, lLinhaI + 3.4, l_local
    
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.FontName = "Roman 10cpi"
'    ImprimeTexto "1", lColunaI + 10.8, lColunaI + 12.05, lLinhaI + 0.7, l_local
    ImprimeCentralizado "2100", lColunaI + 15.1, lColunaI + 19.4, lLinhaI + 0.7, l_local
    If Val(Mid(l_ano_mes, 5, 2)) <= 12 Then
        ImprimeCentralizado Mid(l_ano_mes, 5, 2) & "/" & Mid(l_ano_mes, 1, 4), lColunaI + 15.1, lColunaI + 19.4, lLinhaI + 1.3, l_local
    ElseIf Val(Mid(l_ano_mes, 5, 2)) = 13 Then
        ImprimeCentralizado "11/" & Mid(l_ano_mes, 1, 4), lColunaI + 15.1, lColunaI + 19.4, lLinhaI + 1.3, l_local
    ElseIf Val(Mid(l_ano_mes, 5, 2)) = 14 Then
        ImprimeCentralizado "12/" & Mid(l_ano_mes, 1, 4), lColunaI + 15.1, lColunaI + 19.4, lLinhaI + 1.3, l_local
    End If
    
    
    x_empresa = lTotalSegurado + Format((lTotalEmpregado * 23 / 100), "00,000,000.00") + Format((lValorAutonomo * 15 / 100), "00,000,000.00") - lSalarioFamilia
    ImprimeValor Format(x_empresa, "##,###,##0.00"), lColunaI + 16, lColunaI + 19, lLinhaI + 2.6, l_local
    x_terceiro = Format((lTotalEmpregado * 5.8 / 100), "00,000,000.00")
    ImprimeValor Format(x_terceiro, "##,###,##0.00"), lColunaI + 16, lColunaI + 19, lLinhaI + 4.7, l_local
    'ImprimeValor Format(lTotalSegurado, "##,###,##0.00"), lColunaI + 16, lColunaI + 19.5, lLinhaI + 3.6, l_local
    'x_empresa = CCur(Format((lValorAutonomo * 15 / 100), "00,000,000.00")) + CCur(Format((lTotalEmpregado * 23 / 100), "00,000,000.00"))
    'ImprimeValor Format(x_empresa, "##,###,##0.00"), lColunaI + 16, lColunaI + 19.5, lLinhaI + 4.5, l_local
    'ImprimeCentralizado "0115", lColunaI + 14, lColunaI + 15.5, lLinhaI + 5.4, l_local
    'x_terceiro = Format((lTotalEmpregado * 5.8 / 100), "00,000,000.00")
    'ImprimeValor Format(x_terceiro, "##,###,##0.00"), lColunaI + 16, lColunaI + 19.5, lLinhaI + 5.4, l_local
    'If lSalarioFamilia > 0 Then
    '    ImprimeValor Format(lSalarioFamilia, "##,###,##0.00"), lColunaI + 16, lColunaI + 19.5, lLinhaI + 7.8, l_local
    'End If
    x_total = x_empresa + x_terceiro
    ImprimeValor Format(x_total, "##,###,##0.00"), lColunaI + 16, lColunaI + 19, lLinhaI + 6.4, l_local
    Printer.FontBold = False
    
    'Printer.DrawWidth = 4
    'Printer.Line (lColunaI + 3, lLinhaI + 0.5)-(lColunaI + 3.5, lLinhaI + 0.5)
    'Printer.Line (lColunaI + 3, lLinhaI + 0.5)-(lColunaI + 3, lLinhaI + 1)
    'Printer.Line (lColunaI + 9.3, lLinhaI + 0.5)-(lColunaI + 9.8, lLinhaI + 0.5)
    'Printer.Line (lColunaI + 9.8, lLinhaI + 0.5)-(lColunaI + 9.8, lLinhaI + 1)
    'Printer.Line (lColunaI + 3, lLinhaI + 4)-(lColunaI + 3.5, lLinhaI + 4)
    'Printer.Line (lColunaI + 3, lLinhaI + 3.5)-(lColunaI + 3, lLinhaI + 4)
    'Printer.Line (lColunaI + 9.3, lLinhaI + 4)-(lColunaI + 9.8, lLinhaI + 4)
    'Printer.Line (lColunaI + 9.8, lLinhaI + 3.5)-(lColunaI + 9.8, lLinhaI + 4)
    'Printer.DrawWidth = 4
    
    Printer.FontSize = 12
    Printer.FontBold = True
    ImprimeCentralizado "259-5165", lColunaI + 9, lColunaI + 11.8, lLinhaI + 2.6, l_local
    ImprimeTexto tbl_empresa!Nome, lColunaI + 1.7, lColunaI + 11.8, lLinhaI + 3, l_local
    ImprimeTexto tbl_empresa!Endereco, lColunaI + 1.7, lColunaI + 11.8, lLinhaI + 3.5, l_local
    ImprimeTexto Trim(tbl_empresa!Bairro) & ", " & Trim(tbl_empresa!Cidade) & " - " & Trim(tbl_empresa!Estado), lColunaI + 1.7, lColunaI + 11.8, lLinhaI + 3.9, l_local
    'ImprimeCentralizado Mid(tbl_empresa!CEP, 1, 2) & "." & Mid(tbl_empresa!CEP, 3, 3) & "-" & Mid(tbl_empresa!CEP, 6, 3), lColunaI + 0, lColunaI + 3.1, lLinhaI + 8.2, l_local
    'ImprimeTexto tbl_empresa!Cidade, lColunaI + 3.4, lColunaI + 9, lLinhaI + 7.8, l_local
    'ImprimeCentralizado tbl_empresa!Estado, lColunaI + 8.9, lColunaI + 10.2, lLinhaI + 7.9, l_local
    
    ImprimeTexto Format(lQtdEmpregado, "00"), lColunaI + 1.8, lColunaI + 2.5, lLinhaI + 7, l_local
    ImprimeTexto "Empregados:", lColunaI + 2.5, lColunaI + 6, lLinhaI + 7, l_local
    ImprimeValor Format(lTotalEmpregado, "##,###,##0.00"), lColunaI + 6.5, lColunaI + 8, lLinhaI + 7, l_local
    ImprimeTexto "Empregadores:", lColunaI + 1.8, lColunaI + 6, lLinhaI + 7.5, l_local
    ImprimeValor Format(lValorAutonomo, "##,###,##0.00"), lColunaI + 6.5, lColunaI + 8, lLinhaI + 7.5, l_local
    ImprimeTexto "Salário Família:", lColunaI + 1.8, lColunaI + 6, lLinhaI + 8, l_local
    ImprimeValor Format(lSalarioFamilia, "##,###,##0.00"), lColunaI + 6.5, lColunaI + 8, lLinhaI + 8, l_local
    'ImprimeValor "202.110-2", lColunaI + 5, lColunaI + 9, lLinhaI + 11.5, l_local
    'ImprimeTexto "CNAE: 5050-4/00", lColunaI + 0.2, lColunaI + 6, lLinhaI + 12.1, l_local
    Printer.EndDoc
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    If CalculaValorEmpregado > 0 Then
        lValorAutonomo = fValidaValor2(txt_valor)
        ImpGPS
    End If
    cmd_sair.SetFocus
End Sub
Function RetiraString(x_string As String, numero As Integer) As Integer
    RetiraString = 0
    Dim x_index As Integer
    Dim x_inicio As Integer
    Dim x_numero As Integer
    x_inicio = 1
    x_numero = 1
    If Len(x_string) > 0 Then
        Do Until x_index > Len(x_string)
            x_index = x_index + 1
            If Mid(x_string, x_index, 1) = "@" Then
                If x_numero = numero Then
                    RetiraString = Mid(x_string, x_inicio, x_index - x_inicio)
                    Exit Function
                End If
                x_index = x_index + 2
                x_numero = x_numero + 1
                x_inicio = x_index + 1
            End If
        Loop
    End If
End Function
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate("01/" & msk_data) Then
        MsgBox "Informe o mês/ano.", 64, "Atenção!"
        msk_data.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub ZeraVariaveis()
    lTotalEmpregado = 0
    lQtdEmpregado = 0
    lTotalSegurado = 0
    lSalarioFamilia = 0
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If msk_data.Text = "__/____" Then
        msk_data.Text = Format(g_data_def, "mm") & "/" & Format(g_data_def, "yyyy")
        msk_data.SetFocus
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_imprimir_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    Set tbl_empresa = bd_sgp.OpenTable("Empresas")
    Set tbl_funcionario = bd_sgp.OpenTable("Funcionario")
    Set tbl_movimento_falta_funcionario = bd_sgp.OpenRecordset("Movimento_Falta_Funcionario", dbOpenTable)
    Set tbl_movimento_folha = bd_sgp.OpenTable("Movimento_Folha")
    Set tbl_tabela_folha = bd_sgp.OpenTable("Tabela_Folha")
    Set tbl_tabela_provento_desconto = bd_sgp.OpenTable("Tabela_Provento_Desconto")
    tbl_empresa.Index = "id_codigo"
    tbl_funcionario.Index = "id_codigo"
    tbl_movimento_falta_funcionario.Index = "id_funcionario"
    tbl_movimento_folha.Index = "id_data"
    tbl_tabela_folha.Index = "id_mes_ano"
    tbl_tabela_provento_desconto.Index = "id_codigo"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor.SetFocus
    End If
End Sub
Private Sub msk_data_LostFocus()
    tbl_tabela_folha.Seek "=", Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2)
    If tbl_tabela_folha.NoMatch Then
        MsgBox "Tabela da folha não cadastrada.", 64, "Erro de Consistência!"
        msk_data.SetFocus
    Else
        l_ano_mes = tbl_tabela_folha![Mes Ano]
    End If
End Sub
Private Sub txt_valor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub txt_valor_LostFocus()
    txt_valor = Format(txt_valor, "###,##0.00")
End Sub
