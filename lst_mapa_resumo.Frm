VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form emissao_mapa_resumo 
   Caption         =   "Emissão do Mapa Resumo E.C.F."
   ClientHeight    =   2385
   ClientLeft      =   1965
   ClientTop       =   3705
   ClientWidth     =   4185
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_mapa_resumo.frx":0000
   ScaleHeight     =   2385
   ScaleWidth      =   4185
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   840
      Picture         =   "lst_mapa_resumo.frx":0046
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprime o Recibo."
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2460
      Picture         =   "lst_mapa_resumo.frx":1320
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1320
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3915
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_mapa_resumo.frx":25FA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   360
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_mapa_resumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Início de variáveis padrão para relatório
Dim lColuna(0 To 1) As Currency
Dim lColunaI As Currency
Dim lLinhaI As Currency
Dim lLinhaTab As Currency
Dim lLocal As Integer
Private MovMapaResumo As cMovimentoMapaResumo
Private Sub ZeraVariaveis()
    lLocal = 1
    lLinhaI = 0
    lColunaI = 0
    
    lColuna(0) = lColunaI + 0
    lColuna(1) = lColunaI + 20
    lLinhaTab = 0
End Sub
Private Sub Finaliza()
    'tbl_funcionario.Close
End Sub
Private Sub ImpDados()
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    Printer.FontName = "Arial"
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    Printer.FontBold = True
    'teste para imprimir letra correta
    Printer.FontBold = False
    ImprimeTexto "  ", lColuna(0), lColuna(1), lLinhaTab, lLocal
    Printer.FontBold = True
    'imprime detalhe
    Printer.FontSize = 14
    Printer.DrawWidth = 6
    lLinhaTab = lLinhaI + 2
    ImprimeTexto "R$  (" & Format(txt_valor, "###,##0.00") & ")", lColunaI + 15, lColuna(1), lLinhaTab, lLocal
    Printer.FontSize = 10
    Printer.DrawWidth = 4
    lLinhaTab = lLinhaI + 3
    ImprimeTexto "Recebi da empresa, ", lColunaI + 3, lColuna(1), lLinhaTab, lLocal
    Printer.FontSize = 14
    Printer.DrawWidth = 6
    lLinhaTab = lLinhaI + 2.9
    ImprimeTexto g_nome_empresa, lColunaI + 6.5, lColuna(1), lLinhaTab, lLocal
    Printer.FontSize = 10
    Printer.DrawWidth = 4
    lLinhaTab = lLinhaI + 4
    ImprimeTexto "A quantia supra de ", lColunaI + 1, lColuna(1), lLinhaTab, lLocal
    ImprimeTexto "(" & FazExtenso(fValidaValor2(txt_valor)) & ")", lColunaI + 4.4, lColuna(1), lLinhaTab, lLocal
    lLinhaTab = lLinhaI + 5
    ImprimeTexto "Proveniente a ", lColunaI + 1, lColuna(1), lLinhaTab, lLocal
    ImprimeTexto txt_proveniente & ".", lColunaI + 3.55, lColuna(1), lLinhaTab, lLocal
    lLinhaTab = lLinhaI + 6
    ImprimeTexto g_cidade_empresa & ", " & Day(msk_data) & " de " & Format(msk_data, "mmmm") & " de " & Year(msk_data) & ".", lColunaI + 6, lColuna(1), lLinhaTab, lLocal
    lLinhaTab = lLinhaI + 8.6
    ImprimeCentralizado "___________________________________________________", lColuna(0), lColuna(1), lLinhaTab, lLocal
    lLinhaTab = lLinhaI + 9
    ImprimeCentralizado txt_funcionario, lColuna(0), lColuna(1), lLinhaTab, lLocal
    ImpGrade
    Printer.EndDoc
End Sub
Private Sub ImpGrade()
    Printer.DrawWidth = 2
    
    
    Printer.ForeColor = RGB(0, 0, 256) 'dados em azul
    'Bordas Externas
    Printer.Line (lColunaI + 0, lLinhaI + 0)-(lColunaI + 26, lLinhaI + 0)
    Printer.Line (lColunaI + 0, lLinhaI + 0)-(lColunaI + 0, lLinhaI + 8.5)
    Printer.Line (lColunaI + 26, lLinhaI + 0)-(lColunaI + 26, lLinhaI + 8.5)
    Printer.Line (lColunaI + 0, lLinhaI + 8.5)-(lColunaI + 26, lLinhaI + 8.5)
    
    'Linhas Horizontais do Cabecalho
    Printer.Line (lColunaI + 0, lLinhaI + 1)-(lColunaI + 26, lLinhaI + 1)
    Printer.Line (lColunaI + 0, lLinhaI + 2)-(lColunaI + 26, lLinhaI + 2)
    Printer.Line (lColunaI + 0, lLinhaI + 3)-(lColunaI + 26, lLinhaI + 3)
    
    'Linhas Verticais do Cabecalho
    Printer.Line (lColunaI + 11, lLinhaI + 2)-(lColunaI + 11, lLinhaI + 3)
    Printer.Line (lColunaI + 17, lLinhaI + 0)-(lColunaI + 17, lLinhaI + 1)
    Printer.Line (lColunaI + 19, lLinhaI + 2)-(lColunaI + 19, lLinhaI + 3)
    Printer.Line (lColunaI + 20.1, lLinhaI + 1)-(lColunaI + 20.1, lLinhaI + 3)
    Printer.Line (lColunaI + 22, lLinhaI + 0)-(lColunaI + 22, lLinhaI + 1)
    
    'Linhas do Detalhe
    Printer.Line (lColunaI + 1, lLinhaI + 3.5)-(lColunaI + 9.3, lLinhaI + 3.5)
    Printer.Line (lColunaI + 13.5, lLinhaI + 3.5)-(lColunaI + 24.48, lLinhaI + 3.5)
    Printer.Line (lColunaI + 0, lLinhaI + 4)-(lColunaI + 26, lLinhaI + 4)
    Printer.Line (lColunaI + 0, lLinhaI + 4.5)-(lColunaI + 26, lLinhaI + 4.5)
    Printer.Line (lColunaI + 0, lLinhaI + 5)-(lColunaI + 26, lLinhaI + 5)
    Printer.Line (lColunaI + 1, lLinhaI + 3)-(lColunaI + 1, lLinhaI + 4.5)
    Printer.Line (lColunaI + 2.4, lLinhaI + 3.5)-(lColunaI + 2.4, lLinhaI + 4.5)
    Printer.Line (lColunaI + 3.9, lLinhaI + 3)-(lColunaI + 3.9, lLinhaI + 5)
    Printer.Line (lColunaI + 6.6, lLinhaI + 3.5)-(lColunaI + 6.6, lLinhaI + 5)
    Printer.Line (lColunaI + 9.3, lLinhaI + 3)-(lColunaI + 9.3, lLinhaI + 5)
    Printer.Line (lColunaI + 13.5, lLinhaI + 3)-(lColunaI + 13.5, lLinhaI + 5)
    Printer.Line (lColunaI + 15.33, lLinhaI + 3.5)-(lColunaI + 15.33, lLinhaI + 5)
    Printer.Line (lColunaI + 17.16, lLinhaI + 3.5)-(lColunaI + 17.16, lLinhaI + 5)
    Printer.Line (lColunaI + 18.99, lLinhaI + 3.5)-(lColunaI + 18.99, lLinhaI + 5)
    Printer.Line (lColunaI + 20.82, lLinhaI + 3.5)-(lColunaI + 20.82, lLinhaI + 5)
    Printer.Line (lColunaI + 22.65, lLinhaI + 3.5)-(lColunaI + 22.65, lLinhaI + 5)
    Printer.Line (lColunaI + 24.98, lLinhaI + 3)-(lColunaI + 24.98, lLinhaI + 5)
    
    
    
    Printer.FontSize = 20
    Printer.DrawWidth = 8
    Printer.FontBold = True
    ImprimeCentralizado "teste", lColunaI + 0, lColunaI + 20, lLinhaI + 1, lLocal
    Printer.DrawWidth = 1
End Sub
Private Sub PosicionaProximoRegistro()
    'dta_funcionario.Recordset.FindFirst "Codigo = " & Val(dbcbo_funcionario.BoundText)
    'If Not dta_funcionario.Recordset.EOF Then
    '    dta_funcionario.Recordset.MoveNext
    '    If Not dta_funcionario.Recordset.EOF Then
    '        dbcbo_funcionario.BoundText = dta_funcionario.Recordset!Codigo
    '    Else
    '        dbcbo_funcionario.BoundText = ""
    '    End If
    'End If
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    ImpGrade
    cmd_novo.SetFocus
End Sub
Private Sub ImprimeDadosFormulario()
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    Printer.FontName = "Arial"
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    Printer.FontSize = 14
    Printer.DrawWidth = 6
    ImprimeCentralizado Format(txt_valor, "###,##0.00"), 12.1, 18.1, 2.5, 1
    
    Printer.FontSize = 14
    Printer.DrawWidth = 6
    ImprimeTexto g_nome_empresa, 4, 18, 3.7, 1
    
    Printer.FontSize = 10
    Printer.DrawWidth = 4
    ImprimeTexto FazExtenso(fValidaValor2(txt_valor)), 4, 18, 4.9, 1
    
    ImprimeTexto txt_proveniente, 4, 18, 7.1, 1
    
    ImprimeCentralizado Trim(g_cidade_empresa) & ", " & Day(msk_data) & " de " & Format(msk_data, "mmmm") & " de " & Format(msk_data, "yyyy") & ".", 10, 18, 11.1, 1
    
    ImprimeCentralizado txt_funcionario, 10, 18, 12.5, 1
    Printer.EndDoc
End Sub
Private Sub chk_formulario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = RetiraGString(1)
    'txt_valor.SetFocus
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    If ValidaCampos Then
        If SelecionaImpressoraHP(Me) Then
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
'Private Sub dbcbo_funcionario_GotFocus()
'    txt_funcionario.Visible = False
'End Sub
'Private Sub dbcbo_funcionario_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        KeyCode = 0
'        txt_proveniente.SetFocus
'    End If
'End Sub
'Private Sub dbcbo_funcionario_LostFocus()
'    txt_funcionario.Visible = True
'    If dbcbo_funcionario.BoundText <> "" Then
'        txt_funcionario = dbcbo_funcionario
'        txt_proveniente.SetFocus
'    Else
'        txt_funcionario = ""
'        txt_funcionario.SetFocus
'    End If
'End Sub
'Private Sub Form_Activate()
'    TabelaFuncionarioRefresh
'    If Not IsDate(msk_data) Then
'        msk_data = Format(g_data_def, "dd/mm/yyyy")
'        txt_valor.SetFocus
'    End If
'    Screen.MousePointer = 1
'End Sub
'Private Sub TabelaFuncionarioRefresh()
'    dta_funcionario.RecordSource = "Select * From Funcionario Where Empresa = " & g_empresa & " And Situacao = 'A' Order By Nome"
'    dta_funcionario.Refresh
'End Sub
Private Sub Form_Load()
    CentraForm Me
    Set MovMapaResumo = New cMovimentoMapaResumo
    Set MovMapaResumo.Conexao = Conectar.Conexao
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
