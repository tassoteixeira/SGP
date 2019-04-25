VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form lst_bordero_deposito_avista 
   Caption         =   "Emissão do Borderô de Depósito (Avista)"
   ClientHeight    =   2295
   ClientLeft      =   1965
   ClientTop       =   3705
   ClientWidth     =   6795
   Icon            =   "lst_bordero_deposito_avista.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_bordero_deposito_avista.frx":030A
   ScaleHeight     =   2295
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   1740
      Picture         =   "lst_bordero_deposito_avista.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Imprime borderô de depósito (Avista)."
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4260
      Picture         =   "lst_bordero_deposito_avista.frx":195A
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
      Width           =   6555
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_bordero_deposito_avista.frx":2FEC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_bordero_deposito_avista.frx":42C6
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "lst_bordero_deposito_avista.frx":55A0
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4860
         TabIndex        =   9
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
         TabIndex        =   6
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
      Begin VB.CheckBox chkTodasEmpresas 
         Caption         =   "Imprime &Todas as empresas"
         Height          =   315
         Left            =   3840
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   315
         Index           =   0
         Left            =   3840
         TabIndex        =   8
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   180
      Top             =   1740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_bordero_deposito_avista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Início de variáveis padrão para relatório
Dim lColunaI As Currency
Dim lLinhaI As Currency
Dim lVia As Integer
Dim lLocal As Integer
Dim lTotal As Currency
Dim lTotalQtd As Currency
Dim lBordero As Integer
Dim tbl_empresa As Table
Dim tbl_movimento_cheque_avista As Table
Private Sub ZeraVariaveis()
    lLocal = 1
    lColunaI = 0
    lLinhaI = 2
    lVia = 0
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_empresa.Close
    tbl_movimento_cheque_avista.Close
End Sub
Private Sub ImpDados(x_empresa As Integer)
    Dim x_nome_empresa As String
    Dim x_numero_conta As String
    
    x_nome_empresa = ""
    tbl_empresa.Seek "=", x_empresa
    If Not tbl_empresa.NoMatch Then
        x_nome_empresa = tbl_empresa!Nome
    End If
    If x_empresa = 2 Then
        x_numero_conta = "0051349"
    ElseIf x_empresa = 3 Then
        x_numero_conta = "0051330"
    ElseIf x_empresa = 4 Then
        x_numero_conta = "0051420"
    ElseIf x_empresa = 6 Then
        x_numero_conta = "0051357"
    'teste
        x_nome_empresa = "Rivelino e Ribeiro Ltda."
    'teste
    ElseIf x_empresa = 9 Then
        x_numero_conta = "0054267"
    ElseIf x_empresa = 10 Then
        x_numero_conta = "0055778"
    ElseIf x_empresa = 11 Then
        x_numero_conta = "0012602"
    Else
        x_numero_conta = "0051349"
    End If
    
    lLinhaI = lVia * 7.65 - 7.65 + 1 * lVia
    lColunaI = 1
    
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    Printer.FontName = "Arial"
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    Printer.FontBold = True
    'teste para imprimir letra correta
    Printer.FontBold = False
    ImprimeTexto "  ", lColunaI + 1, lColunaI + 2, lLinhaI + 1, lLocal
    
    Printer.FontSize = 18
    Printer.FontBold = True
    'ImprimeTexto Format(lBordero, "000000"), lColunaI + 15.5, lColunaI + 18.5, lLinhaI + 0.25, lLocal
    
    Printer.FontBold = True
    Printer.FontSize = 10
    Printer.DrawWidth = 4
    If x_empresa = 11 Then
        ImprimeCentralizado "0638", lColunaI + 7.9, lColunaI + 10.7, lLinhaI + 1.4, lLocal
        ImprimeCentralizado "6", lColunaI + 10.7, lColunaI + 11.72, lLinhaI + 1.4, lLocal
    Else
        ImprimeCentralizado "0638", lColunaI + 7.9, lColunaI + 10.7, lLinhaI + 1.4, lLocal
        ImprimeCentralizado "6", lColunaI + 10.7, lColunaI + 11.72, lLinhaI + 1.4, lLocal
    End If
    ImprimeCentralizado Mid(x_numero_conta, 1, 6), lColunaI + 11.72, lColunaI + 16.58, lLinhaI + 1.4, lLocal
    ImprimeCentralizado Mid(x_numero_conta, 7, 1), lColunaI + 16.58, lColunaI + 17.6, lLinhaI + 1.4, lLocal
    ImprimeCentralizado x_nome_empresa, lColunaI + 7.9, lColunaI + 17.6, lLinhaI + 2.2, lLocal
    
    Printer.FontSize = 10
    'ImprimeValor Format(lTotalQtd, "##,##0") & "  ", lColunaI + 0.5, lColunaI + 2.8, lLinhaI + 2.7, lLocal
    ImprimeValor Format(lTotal, "###,###,##0.00") & "  ", lColunaI + 11.72, lColunaI + 16.58, lLinhaI + 4.4, lLocal
    ImprimeValor Format(lTotal, "###,###,##0.00") & "  ", lColunaI + 11.72, lColunaI + 16.58, lLinhaI + 4.9, lLocal
    Call ImpGrade(lVia)
    If lVia = 2 Then
        Printer.EndDoc
    End If
End Sub
Private Sub ImpGrade(x_via As Integer)
    lLinhaI = x_via * 7.65 - 7.65 + 1 * x_via
    lColunaI = 1
    Printer.DrawWidth = 2
    Printer.Line (lColunaI, lLinhaI)-(lColunaI + 17.6, lLinhaI)
    Printer.Line (lColunaI, lLinhaI)-(lColunaI, lLinhaI + 7.65)
    Printer.Line (lColunaI + 17.6, lLinhaI)-(lColunaI + 17.6, lLinhaI + 7.65)
    Printer.Line (lColunaI, lLinhaI + 7.65)-(lColunaI + 17.6, lLinhaI + 7.65)
    
    Printer.DrawWidth = 2
    Printer.FontSize = 16
    Printer.DrawWidth = 8
    Printer.FontBold = True
    ImprimeTexto "BRADESCO", lColunaI + 0.42, lColunaI + 5, lLinhaI + 0.05, lLocal
    
'    Printer.DrawWidth = 1
'    Printer.Line (lColunaI + 0.5, lLinhaI + 0.75)-(lColunaI + 4.03, lLinhaI + 0.75)
'    Printer.Line (lColunaI + 0.5, lLinhaI + 0.95)-(lColunaI + 4.03, lLinhaI + 0.95)
    Printer.FontSize = 6
    Printer.FontBold = True
    ImprimeTexto "CGC 60.746.948", lColunaI + 0.42, lColunaI + 4, lLinhaI + 0.63, lLocal
    
    Printer.DrawWidth = 1
    Printer.FontSize = 8
    Printer.FontBold = False
    ImprimeTexto "COMPROVANTE DE DEPÓSITO / TRANSFERÊNCIA", lColunaI + 3.9, lColunaI + 11.5, lLinhaI + 0.25, lLocal
    Printer.FontSize = 6
    ImprimeTexto "CONTA CORRENTE", lColunaI + 12, lColunaI + 14.5, lLinhaI + 0.25, lLocal
    ImprimeTexto "CONTA DE POUPANÇA", lColunaI + 15, lColunaI + 17.5, lLinhaI + 0.25, lLocal
    Printer.DrawWidth = 4
    
    Printer.Line (lColunaI + 11.5, lLinhaI + 0.2)-(lColunaI + 11.5, lLinhaI + 0.55)
    Printer.Line (lColunaI + 12, lLinhaI + 0.2)-(lColunaI + 12, lLinhaI + 0.55)
    Printer.Line (lColunaI + 11.5, lLinhaI + 0.2)-(lColunaI + 12, lLinhaI + 0.2)
    Printer.Line (lColunaI + 11.5, lLinhaI + 0.55)-(lColunaI + 12, lLinhaI + 0.55)
    Printer.DrawWidth = 8
    Printer.Line (lColunaI + 11.5, lLinhaI + 0.2)-(lColunaI + 12, lLinhaI + 0.55)
    Printer.Line (lColunaI + 11.5, lLinhaI + 0.55)-(lColunaI + 12, lLinhaI + 0.2)
    
    Printer.DrawWidth = 4
    Printer.Line (lColunaI + 14.4, lLinhaI + 0.2)-(lColunaI + 14.4, lLinhaI + 0.55)
    Printer.Line (lColunaI + 14.9, lLinhaI + 0.2)-(lColunaI + 14.9, lLinhaI + 0.55)
    Printer.Line (lColunaI + 14.4, lLinhaI + 0.2)-(lColunaI + 14.9, lLinhaI + 0.2)
    Printer.Line (lColunaI + 14.4, lLinhaI + 0.55)-(lColunaI + 14.9, lLinhaI + 0.55)
    
    
    
    Printer.Line (lColunaI + 0.5, lLinhaI + 1)-(lColunaI + 17.6, lLinhaI + 1)
    Printer.Line (lColunaI + 0.5, lLinhaI + 1.4)-(lColunaI + 7.9, lLinhaI + 1.4)
    Printer.Line (lColunaI + 0.5, lLinhaI + 1.85)-(lColunaI + 17.6, lLinhaI + 1.85)
    Printer.Line (lColunaI + 0.5, lLinhaI + 2.25)-(lColunaI + 7.9, lLinhaI + 2.25)
    Printer.Line (lColunaI + 0.5, lLinhaI + 2.7)-(lColunaI + 17.6, lLinhaI + 2.7)
    Printer.Line (lColunaI + 0.5, lLinhaI + 3.1)-(lColunaI + 7.9, lLinhaI + 3.1)
    Printer.Line (lColunaI + 0.5, lLinhaI + 3.55)-(lColunaI + 17.6, lLinhaI + 3.55)
    Printer.Line (lColunaI + 0.5, lLinhaI + 3.95)-(lColunaI + 17.6, lLinhaI + 3.95)
    Printer.Line (lColunaI + 0.5, lLinhaI + 4.4)-(lColunaI + 17.6, lLinhaI + 4.4)
    Printer.Line (lColunaI + 0.5, lLinhaI + 4.8)-(lColunaI + 17.6, lLinhaI + 4.8)
    Printer.Line (lColunaI + 0.5, lLinhaI + 5.35)-(lColunaI + 17.6, lLinhaI + 5.35)
    
    Printer.Line (lColunaI + 0.5, lLinhaI + 1)-(lColunaI + 0.5, lLinhaI + 5.35)
    Printer.Line (lColunaI + 1.5, lLinhaI + 1.4)-(lColunaI + 1.5, lLinhaI + 4.8)
    Printer.Line (lColunaI + 3.03, lLinhaI + 1.4)-(lColunaI + 3.03, lLinhaI + 4.8)
    Printer.Line (lColunaI + 4.82, lLinhaI + 1.4)-(lColunaI + 4.82, lLinhaI + 5.35)
    Printer.Line (lColunaI + 7.9, lLinhaI + 1)-(lColunaI + 7.9, lLinhaI + 5.35)
    Printer.Line (lColunaI + 11.8, lLinhaI + 3.55)-(lColunaI + 11.8, lLinhaI + 5.35)
    Printer.Line (lColunaI + 17.6, lLinhaI + 1)-(lColunaI + 17.6, lLinhaI + 5.35)
    
    Printer.Line (lColunaI + 2.05, lLinhaI + 5.35)-(lColunaI + 2.05, lLinhaI + 5.65)
    Printer.Line (lColunaI + 3.58, lLinhaI + 5.35)-(lColunaI + 3.58, lLinhaI + 5.65)
    Printer.Line (lColunaI + 4.59, lLinhaI + 5.35)-(lColunaI + 4.59, lLinhaI + 5.65)
    Printer.Line (lColunaI + 6.6, lLinhaI + 5.35)-(lColunaI + 6.6, lLinhaI + 5.65)
    Printer.Line (lColunaI + 10.45, lLinhaI + 5.35)-(lColunaI + 10.45, lLinhaI + 6.1)
    Printer.Line (lColunaI + 0.5, lLinhaI + 5.65)-(lColunaI + 17.6, lLinhaI + 5.65)
    
    Printer.Line (lColunaI + 10.7, lLinhaI + 1.3)-(lColunaI + 10.7, lLinhaI + 1.85)
    Printer.Line (lColunaI + 11.72, lLinhaI + 1)-(lColunaI + 11.72, lLinhaI + 1.85)
    Printer.Line (lColunaI + 16.58, lLinhaI + 1.3)-(lColunaI + 16.58, lLinhaI + 1.85)
    
    
    Printer.FontSize = 6
    Printer.FontBold = False
    ImprimeCentralizado "RELACIONAR SOMENTE CHEQUES SUPERIORES (VIDE-VERSO)", lColunaI + 0.6, lColunaI + 7.8, lLinhaI + 1.1, lLocal
    ImprimeCentralizado "CÓDIGO AGÊNCIA", lColunaI + 7.9, lColunaI + 10.7, lLinhaI + 1.1, lLocal
    ImprimeCentralizado "DÍG.", lColunaI + 10.7, lColunaI + 11.72, lLinhaI + 1.1, lLocal
    ImprimeCentralizado "NÚMERO DA CONTA", lColunaI + 11.72, lColunaI + 16.58, lLinhaI + 1.1, lLocal
    ImprimeCentralizado "DÍG.", lColunaI + 16.58, lColunaI + 17.6, lLinhaI + 1.1, lLocal
    
    ImprimeCentralizado "BANCO", lColunaI + 0.5, lColunaI + 1.5, lLinhaI + 1.5, lLocal
    ImprimeCentralizado "SÉRIE", lColunaI + 1.5, lColunaI + 3.03, lLinhaI + 1.5, lLocal
    ImprimeCentralizado "N. CHEQUE", lColunaI + 3.03, lColunaI + 4.82, lLinhaI + 1.5, lLocal
    ImprimeCentralizado "VALOR", lColunaI + 4.82, lColunaI + 7.9, lLinhaI + 1.5, lLocal
    
    ImprimeTexto "PARA CRÉDITO DE", lColunaI + 7.9, lColunaI + 17, lLinhaI + 1.9, lLocal
    ImprimeTexto "NOME DO DEPOSITANTE / REMETENTE", lColunaI + 7.9, lColunaI + 17, lLinhaI + 2.75, lLocal
    ImprimeTexto "VALOR EM DINHEIRO", lColunaI + 7.9, lColunaI + 17, lLinhaI + 3.6, lLocal
    ImprimeTexto "VALOR CHEQUES SUPERIORES", lColunaI + 7.9, lColunaI + 17, lLinhaI + 4.1, lLocal
    ImprimeTexto "VALOR CHEQUES INFERIORES", lColunaI + 7.9, lColunaI + 17, lLinhaI + 4.46, lLocal
    ImprimeTexto "TOTAL", lColunaI + 7.9, lColunaI + 17, lLinhaI + 4.9, lLocal
    
    Printer.FontSize = 4
    Printer.FontBold = True
    ImprimeCentralizado "TOTAL A TRANSPORTAR PARA O CAMPO", lColunaI + 0.5, lColunaI + 4.3, lLinhaI + 4.9, lLocal
    ImprimeCentralizado "" & Chr(34) & "VALOR CHEQUES SUPERIORES" & Chr(34) & "", lColunaI + 0.5, lColunaI + 4.3, lLinhaI + 5.1, lLocal
    
    ImprimeCentralizado "AGÊNCIA", lColunaI + 0.5, lColunaI + 2.05, lLinhaI + 5.41, lLocal
    ImprimeCentralizado "N.TERM.N.AUT.", lColunaI + 2.05, lColunaI + 3.58, lLinhaI + 5.41, lLocal
    ImprimeCentralizado "DATA", lColunaI + 3.58, lColunaI + 4.59, lLinhaI + 5.41, lLocal
    ImprimeCentralizado "N. CONTA", lColunaI + 4.59, lColunaI + 6.6, lLinhaI + 5.41, lLocal
    ImprimeCentralizado "VALOR", lColunaI + 6.6, lColunaI + 10.45, lLinhaI + 5.41, lLocal
    ImprimeCentralizado "NOME DO FAVORECIDO", lColunaI + 10.45, lColunaI + 17.6, lLinhaI + 5.41, lLocal
    
    
    'ImprimeTexto "AUTORIZAMOS O CRÉDITO NA AGÊNCIA E CONTA - CORRENTE ACIMA DOS", lColunaI + 10.25, lColunaI + 18.5, lLinhaI + 3.2, lLocal
    'ImprimeTexto "VALORES DOS CHEQUES CONTIDOS NESTE BORDERÔ, APÓS CONFERÊNCIA PELO", lColunaI + 9.3, lColunaI + 18.5, lLinhaI + 3.7, lLocal
    'ImprimeTexto "BANCO, NAS DATAS INDICADAS NO VERSO DE CADA CHEQUE.", lColunaI + 9.3, lColunaI + 18.5, lLinhaI + 4.2, lLocal
    
    
    Printer.DrawWidth = 1
    Printer.FontSize = 6
    Printer.FontBold = False
    ImprimeTexto "CONFIRA O NÚMERO DA CONTA, VALOR E NOME", lColunaI + 6.6, lColunaI + 12, lLinhaI + 6.6, lLocal
End Sub
Private Sub Relatorio()
    Dim i As Integer
    ZeraVariaveis
    For i = 1 To 11
        If chkTodasEmpresas = True Or (chkTodasEmpresas = False And i = g_empresa) Then
            If TotalizaChequeAVista(i) > 0 Then
                lBordero = lBordero + 1
                lVia = 1
                Call ImpDados(i)
                lVia = 2
                Call ImpDados(i)
            End If
        End If
    Next
    cmd_sair.SetFocus
End Sub
Private Sub chkTodasEmpresas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_i.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cmd_imprimir.SetFocus
    Else
        msk_data = RetiraGString(1)
        chkTodasEmpresas.SetFocus
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
    cmd_imprimir.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cmd_imprimir.SetFocus
    Else
        msk_data_i = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    If ValidaCampos Then
        If SelecionaImpressoraHP(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        End If
    End If
End Sub
Function TotalizaChequeAVista(x_empresa As Integer) As Currency
    TotalizaChequeAVista = 0
    lTotal = 0
    lTotalQtd = 0
    With tbl_movimento_cheque_avista
        If .RecordCount > 0 Then
            .Seek ">=", x_empresa, CDate(msk_data_i), " ", " ", 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> x_empresa Or ![Data de Emissao] > CDate(msk_data_f) Then
                        Exit Do
                    End If
                    TotalizaChequeAVista = TotalizaChequeAVista + !valor
                    lTotal = lTotal + !valor
                    lTotalQtd = lTotalQtd + 1
                    .MoveNext
                Loop
            End If
        End If
    End With
End Function
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
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        chkTodasEmpresas.Value = True
        msk_data_i.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        cmd_imprimir.SetFocus
        Screen.MousePointer = 1
    End If
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
    Set tbl_movimento_cheque_avista = bd_sgp.OpenTable("Movimento_Cheque_Avista")
    tbl_empresa.Index = "id_codigo"
    tbl_movimento_cheque_avista.Index = "id_digitacao"
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
        cmd_imprimir.SetFocus
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
        chkTodasEmpresas.SetFocus
    End If
End Sub
