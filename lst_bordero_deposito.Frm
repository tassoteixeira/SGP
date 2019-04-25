VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form lst_bordero_deposito 
   Caption         =   "Emissão do Borderô de Depósito (Custódia)"
   ClientHeight    =   2715
   ClientLeft      =   1965
   ClientTop       =   3705
   ClientWidth     =   6795
   Icon            =   "lst_bordero_deposito.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_bordero_deposito.frx":030A
   ScaleHeight     =   2715
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   1740
      Picture         =   "lst_bordero_deposito.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Imprime borderô de depósito (Custódia)."
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4260
      Picture         =   "lst_bordero_deposito.frx":195A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1740
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6555
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "lst_bordero_deposito.frx":2FEC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_bordero_deposito.frx":42C6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_bordero_deposito.frx":55A0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chkTodasEmpresas 
         Caption         =   "Imprime &Todas as empresas"
         Height          =   315
         Left            =   3840
         TabIndex        =   12
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txt_bordero 
         Height          =   315
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   11
         Top             =   1080
         Width           =   855
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
         TabIndex        =   7
         Top             =   660
         Width           =   975
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
         Caption         =   "&Borderô"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   180
      Top             =   2100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_bordero_deposito"
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

Private MovCheque As New cMovimentoCheque
Private Sub ZeraVariaveis()
    lLocal = 1
    lColunaI = 0
    lLinhaI = 2
    lVia = 0
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_empresa.Close
    Set MovCheque = Nothing
End Sub
Private Sub ImpDados(x_empresa As Integer)
    Dim x_nome_empresa As String
    Dim x_numero_conta As String
    Dim x_extenso As String
    
    x_nome_empresa = ""
    tbl_empresa.Seek "=", x_empresa
    If Not tbl_empresa.NoMatch Then
        x_nome_empresa = tbl_empresa!Nome
    End If
    If x_empresa = 2 Then
        x_numero_conta = "0055948"
    ElseIf x_empresa = 3 Then
        x_numero_conta = "0055956"
    ElseIf x_empresa = 4 Then
        x_numero_conta = "0055980"
    ElseIf x_empresa = 6 Then
        x_numero_conta = "0055964"
    'teste
        x_nome_empresa = "Rivelino e Ribeiro Ltda."
    'teste
    ElseIf x_empresa = 9 Then
        x_numero_conta = "0055972"
    ElseIf x_empresa = 10 Then
        x_numero_conta = "0056111"
    ElseIf x_empresa = 11 Then
        x_numero_conta = "0056138"
    Else
        x_numero_conta = "0055948"
    End If
    x_extenso = FazExtenso(lTotal)
    
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
    ImprimeTexto Format(lBordero, "000000"), lColunaI + 15.5, lColunaI + 18.5, lLinhaI + 0.25, lLocal
    
    Printer.FontBold = True
    Printer.FontSize = 10
    Printer.DrawWidth = 4
    If x_empresa = 11 Then
        ImprimeCentralizado "0638", lColunaI + 0.5, lColunaI + 2.8, lLinhaI + 1.3, lLocal
        ImprimeCentralizado "6", lColunaI + 2.8, lColunaI + 3.5, lLinhaI + 1.3, lLocal
    Else
        ImprimeCentralizado "0638", lColunaI + 0.5, lColunaI + 2.8, lLinhaI + 1.3, lLocal
        ImprimeCentralizado "6", lColunaI + 2.8, lColunaI + 3.5, lLinhaI + 1.3, lLocal
    End If
    ImprimeCentralizado Mid(x_numero_conta, 1, 6), lColunaI + 3.5, lColunaI + 7, lLinhaI + 1.3, lLocal
    ImprimeCentralizado Mid(x_numero_conta, 7, 1), lColunaI + 7, lColunaI + 7.7, lLinhaI + 1.3, lLocal
    ImprimeCentralizado x_nome_empresa, lColunaI + 7.9, lColunaI + 15.6, lLinhaI + 1.3, lLocal
    ImprimeCentralizado msk_data, lColunaI + 15.6, lColunaI + 18, lLinhaI + 1.3, lLocal
    ImprimeCentralizado "HIDROLÂNDIA", lColunaI + 0.5, lColunaI + 7.7, lLinhaI + 2, lLocal
    
    Printer.FontSize = 6
    ImprimeTexto "( " & x_extenso, lColunaI + 7.9, lColunaI + 18, lLinhaI + 2.1, lLocal
    ImprimeValor " )", lColunaI + 17.5, lColunaI + 18, lLinhaI + 2.1, lLocal
    
    Printer.FontSize = 10
    ImprimeValor Format(lTotalQtd, "##,##0") & "  ", lColunaI + 0.5, lColunaI + 2.8, lLinhaI + 2.7, lLocal
    ImprimeValor Format(lTotal, "###,###,##0.00") & "  ", lColunaI + 2.8, lColunaI + 7.7, lLinhaI + 2.7, lLocal
    Call ImpGrade(lVia)
    If lVia = 3 Then
        Printer.EndDoc
    End If
End Sub
Private Sub ImpGrade(x_via As Integer)
    lLinhaI = x_via * 7.65 - 7.65 + 1 * x_via
    lColunaI = 1
    Printer.DrawWidth = 2
    Printer.Line (lColunaI, lLinhaI)-(lColunaI + 18.5, lLinhaI)
    Printer.Line (lColunaI, lLinhaI)-(lColunaI, lLinhaI + 7.65)
    Printer.Line (lColunaI + 18.5, lLinhaI)-(lColunaI + 18.5, lLinhaI + 7.65)
    Printer.Line (lColunaI, lLinhaI + 6.75)-(lColunaI + 18.5, lLinhaI + 6.75)
    Printer.Line (lColunaI, lLinhaI + 7.65)-(lColunaI + 18.5, lLinhaI + 7.65)
    
    Printer.DrawWidth = 2
    Printer.FontSize = 18
    Printer.DrawWidth = 8
    Printer.FontBold = True
    ImprimeTexto "BRADESCO", lColunaI + 0.42, lColunaI + 5, lLinhaI + 0.05, lLocal
    
    Printer.DrawWidth = 1
    Printer.Line (lColunaI + 0.5, lLinhaI + 0.75)-(lColunaI + 4.03, lLinhaI + 0.75)
    Printer.Line (lColunaI + 0.5, lLinhaI + 0.95)-(lColunaI + 4.03, lLinhaI + 0.95)
    Printer.FontSize = 4
    Printer.FontBold = True
    ImprimeTexto "BANCO BRADESCO S.A.", lColunaI + 0.42, lColunaI + 4, lLinhaI + 0.76, lLocal
    
    Printer.DrawWidth = 1
    Printer.FontSize = 10
    Printer.FontBold = False
    ImprimeTexto "BORDERÔ PARA ENTREGA DE CHEQUES EM CUSTÓDIA   N.", lColunaI + 4.6, lColunaI + 18.5, lLinhaI + 0.45, lLocal
    Printer.DrawWidth = 4
    Printer.Line (lColunaI + 0.5, lLinhaI + 1)-(lColunaI + 18, lLinhaI + 1)
    Printer.Line (lColunaI + 0.5, lLinhaI + 1.7)-(lColunaI + 18, lLinhaI + 1.7)
    Printer.Line (lColunaI + 0.5, lLinhaI + 2.4)-(lColunaI + 18, lLinhaI + 2.4)
    Printer.Line (lColunaI + 0.5, lLinhaI + 3.1)-(lColunaI + 18, lLinhaI + 3.1)
    Printer.Line (lColunaI + 0.5, lLinhaI + 1)-(lColunaI + 0.5, lLinhaI + 3.1)
    Printer.Line (lColunaI + 2.8, lLinhaI + 1)-(lColunaI + 2.8, lLinhaI + 1.7)
    Printer.Line (lColunaI + 3.5, lLinhaI + 1)-(lColunaI + 3.5, lLinhaI + 1.7)
    Printer.Line (lColunaI + 7, lLinhaI + 1)-(lColunaI + 7, lLinhaI + 1.7)
    Printer.Line (lColunaI + 7.7, lLinhaI + 1)-(lColunaI + 7.7, lLinhaI + 3.1)
    Printer.Line (lColunaI + 15.6, lLinhaI + 1)-(lColunaI + 15.6, lLinhaI + 1.7)
    Printer.Line (lColunaI + 2.8, lLinhaI + 2.4)-(lColunaI + 2.8, lLinhaI + 3.1)
    Printer.Line (lColunaI + 18, lLinhaI + 1)-(lColunaI + 18, lLinhaI + 3.1)
    Printer.FontSize = 6
    Printer.FontBold = False
    ImprimeCentralizado "CÓD.AGÊNCIA", lColunaI + 0.5, lColunaI + 2.8, lLinhaI + 1.02, lLocal
    ImprimeCentralizado "DÍG.", lColunaI + 2.8, lColunaI + 3.5, lLinhaI + 1.02, lLocal
    ImprimeCentralizado "CONTA N.", lColunaI + 3.5, lColunaI + 7, lLinhaI + 1.02, lLocal
    ImprimeCentralizado "DÍG.", lColunaI + 7, lColunaI + 7.7, lLinhaI + 1.02, lLocal
    ImprimeCentralizado "NOME DA EMPRESA", lColunaI + 7.7, lColunaI + 15.5, lLinhaI + 1.02, lLocal
    ImprimeCentralizado "DATA", lColunaI + 15.5, lColunaI + 18, lLinhaI + 1.02, lLocal
    ImprimeCentralizado "NOME DA AGÊNCIA", lColunaI + 0.5, lColunaI + 7.7, lLinhaI + 1.72, lLocal
    ImprimeCentralizado "VALOR POR EXTENSO", lColunaI + 7.7, lColunaI + 18, lLinhaI + 1.72, lLocal
    ImprimeCentralizado "QTDE. CHEQUES", lColunaI + 0.5, lColunaI + 2.8, lLinhaI + 2.42, lLocal
    ImprimeCentralizado "VALOR TOTAL DOS CHEQUES", lColunaI + 2.8, lColunaI + 7.7, lLinhaI + 2.42, lLocal
    
    If x_via = 1 Then
        ImprimeTexto "AUTORIZAMOS O CRÉDITO NA AGÊNCIA E CONTA - CORRENTE ACIMA DOS", lColunaI + 10.25, lColunaI + 18.5, lLinhaI + 3.2, lLocal
        ImprimeTexto "VALORES DOS CHEQUES CONTIDOS NESTE BORDERÔ, APÓS CONFERÊNCIA PELO", lColunaI + 9.3, lColunaI + 18.5, lLinhaI + 3.7, lLocal
        ImprimeTexto "BANCO, NAS DATAS INDICADAS NO VERSO DE CADA CHEQUE.", lColunaI + 9.3, lColunaI + 18.5, lLinhaI + 4.2, lLocal
    ElseIf x_via = 2 Then
        ImprimeTexto "CONFERIMOS E CONFIRMAMOS O VALOR INFORMADO COM OS CHEQUES", lColunaI + 8.6, lColunaI + 18.5, lLinhaI + 3.2, lLocal
        ImprimeTexto "APRESENTADOS.", lColunaI + 8.6, lColunaI + 18.5, lLinhaI + 3.7, lLocal
        ImprimeTexto "O VALOR INFORMADO NÃO CONFERE COM O APURADO DE", lColunaI + 8.6, lColunaI + 18.5, lLinhaI + 4.2, lLocal
        ImprimeTexto "R$", lColunaI + 8.6, lColunaI + 18.5, lLinhaI + 4.7, lLocal
        ImprimeTexto "CONFORME ESCLARECIMENTOS NO VERSO.", lColunaI + 13, lColunaI + 18.5, lLinhaI + 4.7, lLocal
        Printer.DrawWidth = 4
        Printer.Line (lColunaI + 7.7, lLinhaI + 3.2)-(lColunaI + 8.4, lLinhaI + 3.2)
        Printer.Line (lColunaI + 7.7, lLinhaI + 3.9)-(lColunaI + 8.4, lLinhaI + 3.9)
        Printer.Line (lColunaI + 7.7, lLinhaI + 3.2)-(lColunaI + 7.7, lLinhaI + 3.9)
        Printer.Line (lColunaI + 8.4, lLinhaI + 3.2)-(lColunaI + 8.4, lLinhaI + 3.9)
        Printer.Line (lColunaI + 7.7, lLinhaI + 4.2)-(lColunaI + 8.4, lLinhaI + 4.2)
        Printer.Line (lColunaI + 7.7, lLinhaI + 4.9)-(lColunaI + 8.4, lLinhaI + 4.9)
        Printer.Line (lColunaI + 7.7, lLinhaI + 4.2)-(lColunaI + 7.7, lLinhaI + 4.9)
        Printer.Line (lColunaI + 8.4, lLinhaI + 4.2)-(lColunaI + 8.4, lLinhaI + 4.9)
        
        Printer.Line (lColunaI + 9.1, lLinhaI + 4.5)-(lColunaI + 13, lLinhaI + 4.5)
        Printer.Line (lColunaI + 9.1, lLinhaI + 5.2)-(lColunaI + 13, lLinhaI + 5.2)
        Printer.Line (lColunaI + 9.1, lLinhaI + 4.5)-(lColunaI + 9.1, lLinhaI + 5.2)
        Printer.Line (lColunaI + 13, lLinhaI + 4.5)-(lColunaI + 13, lLinhaI + 5.2)
    ElseIf x_via = 3 Then
        ImprimeTexto "OS VALORES E QUANTIDADES DECLARADAS ESTÃO SUJEITOS À CONFERÊNCIA", lColunaI + 9.6, lColunaI + 18.5, lLinhaI + 3.2, lLocal
        ImprimeTexto "E VALIDAÇÃO POSTERIOR.", lColunaI + 9.1, lColunaI + 18.5, lLinhaI + 3.7, lLocal
    End If
    
    Printer.DrawWidth = 4
    Printer.Line (lColunaI + 0.5, lLinhaI + 3.3)-(lColunaI + 1, lLinhaI + 3.3)
    Printer.Line (lColunaI + 6.3, lLinhaI + 3.3)-(lColunaI + 6.8, lLinhaI + 3.3)
    Printer.Line (lColunaI + 0.5, lLinhaI + 3.3)-(lColunaI + 0.5, lLinhaI + 3.85)
    Printer.Line (lColunaI + 6.8, lLinhaI + 3.3)-(lColunaI + 6.8, lLinhaI + 3.85)
    
    Printer.Line (lColunaI + 0.5, lLinhaI + 6.65)-(lColunaI + 1, lLinhaI + 6.65)
    Printer.Line (lColunaI + 6.3, lLinhaI + 6.65)-(lColunaI + 6.8, lLinhaI + 6.65)
    Printer.Line (lColunaI + 0.5, lLinhaI + 6.05)-(lColunaI + 0.5, lLinhaI + 6.65)
    Printer.Line (lColunaI + 6.8, lLinhaI + 6.05)-(lColunaI + 6.8, lLinhaI + 6.65)
    
    Printer.DrawWidth = 4
    Printer.Line (lColunaI + 8, lLinhaI + 6.1)-(lColunaI + 15.5, lLinhaI + 6.1)
    If x_via > 1 Then
        Printer.Line (lColunaI + 16, lLinhaI + 6.1)-(lColunaI + 18, lLinhaI + 6.1)
        Printer.FontBold = True
        Printer.FontSize = 10
        Printer.DrawWidth = 4
        ImprimeTexto "/", lColunaI + 16.55, lColunaI + 18, lLinhaI + 5.7, lLocal
        ImprimeTexto "/", lColunaI + 17.2, lColunaI + 18, lLinhaI + 5.7, lLocal
    End If
    Printer.DrawWidth = 1
    Printer.FontSize = 6
    Printer.FontBold = False
    ImprimeCentralizado "EMPRESA (ASSINATURA, NOME E CARGO)", lColunaI + 8, lColunaI + 15.5, lLinhaI + 6.2, lLocal
    If x_via > 1 Then
        ImprimeCentralizado "DATA", lColunaI + 16, lColunaI + 18, lLinhaI + 6.2, lLocal
    End If
    Printer.FontSize = 4
    If x_via = 1 Then
        ImprimeTexto "1a VIA - CONTROLE DO BANCO", lColunaI + 1, lColunaI + 10, lLinhaI + 7.3, lLocal
    ElseIf x_via = 2 Then
        ImprimeTexto "2a VIA - CONFERÊNCIA DO LOTE/RETORNO A EMPRESA", lColunaI + 1, lColunaI + 10, lLinhaI + 7.3, lLocal
    ElseIf x_via = 3 Then
        ImprimeTexto "3a VIA - PROTOCOLO INICIAL DA EMPRESA", lColunaI + 1, lColunaI + 10, lLinhaI + 7.3, lLocal
    End If
End Sub
Private Sub Relatorio()
    Dim i As Integer
    ZeraVariaveis
    lBordero = CLng(txt_bordero) - 1
    For i = 1 To 11
        If chkTodasEmpresas = True Or (chkTodasEmpresas = False And i = g_empresa) Then
            If TotalizaChequePreDatado(i) > 0 Then
                lBordero = lBordero + 1
                lVia = 1
                Call ImpDados(i)
                lVia = 2
                Call ImpDados(i)
                lVia = 3
                Call ImpDados(i)
            End If
        End If
    Next
    cmd_sair.SetFocus
End Sub
Private Sub chkTodasEmpresas_KeyPress(KeyAscii As Integer)
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
        txt_bordero.SetFocus
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
    txt_bordero.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        txt_bordero.SetFocus
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
Function TotalizaChequePreDatado(x_empresa As Integer) As Currency
    TotalizaChequePreDatado = MovCheque.TotalEmissaoPeriodo(x_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), "1", "9", "0", "P")
    lTotal = lTotal + TotalizaChequePreDatado
    lTotalQtd = MovCheque.TotalQtdEmissaoPeriodo(x_empresa, CDate(msk_data_i.Text), CDate(msk_data_f.Text), "1", "9", "0", "P")
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
        msk_data_i.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        txt_bordero.Text = 1
        chkTodasEmpresas.Value = True
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
    tbl_empresa.Index = "id_codigo"
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
        txt_bordero.SetFocus
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
Private Sub txt_bordero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chkTodasEmpresas.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_bordero_LostFocus()
    txt_bordero = Format(txt_bordero, "000000")
End Sub
