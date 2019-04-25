VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_rpa 
   Caption         =   "Emissão de R.P.A."
   ClientHeight    =   5115
   ClientLeft      =   1965
   ClientTop       =   3705
   ClientWidth     =   6975
   Icon            =   "lst_rpa.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_rpa.frx":030A
   ScaleHeight     =   5115
   ScaleWidth      =   6975
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4320
      Picture         =   "lst_rpa.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4140
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   1860
      Picture         =   "lst_rpa.frx":19E2
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Imprime recibo de RPA."
      Top             =   4140
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   3915
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_rpa.frx":2FEC
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox txt_orgao_emissor 
         Height          =   300
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   14
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txt_identidade 
         Height          =   300
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   12
         Top             =   2040
         Width           =   1515
      End
      Begin VB.TextBox txt_cpf 
         Height          =   300
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1680
         Width           =   1995
      End
      Begin VB.TextBox txt_cidade 
         Height          =   300
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   16
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox txt_razao_social 
         Height          =   300
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   2
         Top             =   240
         Width           =   4875
      End
      Begin VB.TextBox txt_nome 
         Height          =   300
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   21
         Top             =   3480
         Width           =   4875
      End
      Begin VB.TextBox txt_valor 
         Height          =   300
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txt_proveniente 
         Height          =   300
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   6
         Top             =   960
         Width           =   4875
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   300
         Left            =   1680
         TabIndex        =   18
         Top             =   3120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_cgc 
         Height          =   300
         Left            =   1680
         TabIndex        =   4
         Top             =   600
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   18
         Format          =   "__.___.___/____-__"
         Mask            =   "##.###.###/####-##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         Caption         =   "C.G.C."
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   630
         Width           =   1515
      End
      Begin VB.Label Label9 
         Caption         =   "Órgão Emissor"
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   1515
      End
      Begin VB.Label Label8 
         Caption         =   "N. Identidade"
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   1515
      End
      Begin VB.Label Label7 
         Caption         =   "C.P.F."
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "Cidade"
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "Razão social"
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "&Valor"
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Serviços prestados"
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "N&ome"
         Height          =   300
         Left            =   120
         TabIndex        =   20
         Top             =   3480
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &Emissão"
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   3120
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   180
      Top             =   4140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_rpa"
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
Private Sub ZeraVariaveis()
    lColunaI = 0
    lColuna(0) = lColunaI + 0
    lColuna(1) = lColunaI + 20
    lLinhaI = 2
    lLinhaTab = 0
    lLocal = 1
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
End Sub
Private Sub LimpaTela()
    txt_razao_social = g_nome_empresa
    txt_proveniente = "LOCAÇÃO/MANUTENÇÃO NO SISTEMA GERENCIADOR DE POSTO"
    txt_valor = "130,00"
    txt_cpf = "589.766.631-87"
    txt_identidade = "1.770.444"
    txt_orgao_emissor = "SSP-GO"
    txt_cidade = "Goiânia"
    msk_data = Format(g_data_def, "dd/mm/yyyy")
    txt_nome = "TASSO TEIXEIRA"
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    ImprimeDados
    cmd_sair.SetFocus
End Sub
Private Sub ImprimeDados()
    Dim x_extenso As String
    Dim x_extenso_1 As String
    Dim x_extenso_2 As String
    Dim x_tamanho_string As Currency
    Dim i As Integer
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    Printer.FontName = "Arial"
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    
    Printer.FontSize = 14
    Printer.DrawWidth = 6
    ImprimeTexto txt_razao_social, 0.5, 14, 3, 1
    ImprimeTexto msk_cgc, 14.7, 19.5, 3, 1
    Printer.FontSize = 10
    Printer.DrawWidth = 4
    x_extenso_1 = ""
    x_extenso_2 = ""
    x_extenso = txt_proveniente
    For i = 1 To Len(x_extenso)
        x_extenso_1 = Mid(x_extenso, 1, i)
        x_tamanho_string = Printer.TextWidth(x_extenso_1)
        If x_tamanho_string > 6.3 Then
            x_extenso_1 = Mid(x_extenso, 1, i - 1)
            x_extenso_2 = Mid(x_extenso, i, Len(x_extenso))
            Exit For
        End If
    Next
    ImprimeValor x_extenso_1, 13.2, 19.5, 3.8, 1
    ImprimeTexto x_extenso_2, 0.5, 6.5, 4.5, 1
    ImprimeValor txt_valor, 10, 14, 4.5, 1
    x_extenso_1 = ""
    x_extenso_2 = ""
    x_extenso = FazExtenso(fValidaValor(txt_valor))
    For i = 1 To Len(x_extenso)
        x_extenso_1 = Mid(x_extenso, 1, i)
        x_tamanho_string = Printer.TextWidth(x_extenso_1)
        If x_tamanho_string > 4.9 Then
            x_extenso_1 = Mid(x_extenso, 1, i - 1)
            x_extenso_2 = Mid(x_extenso, i, Len(x_extenso))
            Exit For
        End If
    Next
    x_extenso = x_extenso_2
    For i = 1 To 1000
        x_extenso_2 = x_extenso & String(i, "*")
        x_tamanho_string = Printer.TextWidth(x_extenso_2)
        If x_tamanho_string > 19 Then
            x_extenso_2 = x_extenso & String(i - 1, "*")
            Exit For
        End If
    Next
    ImprimeValor x_extenso_1, 14.6, 19.5, 4.5, 1
    ImprimeTexto x_extenso_2, 0.5, 19.5, 5.1, 1
    For i = 1 To 1000
        x_extenso_2 = String(i, "*")
        x_tamanho_string = Printer.TextWidth(x_extenso_2)
        If x_tamanho_string > 19 Then
            x_extenso_2 = String(i - 1, "*")
            Exit For
        End If
    Next
    ImprimeTexto x_extenso_2, 0.5, 14.3, 5.8, 1
    ImprimeValor txt_valor, 16.5, 19, 6.95, 1
    ImprimeValor txt_valor, 16.5, 19, 8.3, 1
    ImprimeValor txt_valor, 16.5, 19, 11.9, 1
    ImprimeCentralizado txt_cpf, 1.7, 9.5, 12.2, 1
    ImprimeCentralizado txt_identidade, 0.4, 4, 13.9, 1
    ImprimeCentralizado txt_orgao_emissor, 4, 9.5, 13.9, 1
    ImprimeCentralizado txt_cidade, 0.4, 6.8, 15, 1
    ImprimeCentralizado Mid(msk_data, 1, 2), 7, 7.75, 14.9, 1
    ImprimeCentralizado Mid(msk_data, 4, 2), 7.75, 8.63, 14.9, 1
    ImprimeCentralizado Mid(msk_data, 7, 4), 8.55, 9.55, 14.9, 1
    ImprimeCentralizado txt_nome, 9.7, 19.5, 15, 1
    Printer.EndDoc
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If txt_razao_social = "" Then
        MsgBox "Informe a razão social.", 64, "Atenção!"
        txt_razao_social.SetFocus
    ElseIf txt_proveniente = "" Then
        MsgBox "Informe os servicos prestados.", 64, "Atenção!"
        txt_proveniente.SetFocus
    ElseIf Not fValidaValor(txt_valor) > 0 Then
        MsgBox "Informe o valor.", 64, "Atenção!"
        txt_valor.SetFocus
    ElseIf txt_cpf = "" Then
        MsgBox "Informe o CPF.", 64, "Atenção!"
        txt_cpf.SetFocus
    ElseIf txt_identidade = "" Then
        MsgBox "Informe o número da identidade.", 64, "Atenção!"
        txt_identidade.SetFocus
    ElseIf txt_orgao_emissor = "" Then
        MsgBox "Informe o órgão emissor.", 64, "Atenção!"
        txt_orgao_emissor.SetFocus
    ElseIf txt_cidade = "" Then
        MsgBox "Informe o nome da cidade.", 64, "Atenção!"
        txt_cidade.SetFocus
    ElseIf Not IsDate(msk_data) Then
        MsgBox "Informe a data de emissão.", 64, "Atenção!"
        msk_data.SetFocus
    ElseIf txt_nome = "" Then
        MsgBox "Informe o nome.", 64, "Atenção!"
        txt_nome.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = RetiraGString(1)
    txt_nome.SetFocus
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
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        LimpaTela
        txt_razao_social.SetFocus
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
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_cgc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_proveniente.SetFocus
    End If
End Sub
Private Sub msk_cgc_LostFocus()
    If Len(msk_cgc.Text) = 14 Then
        If Not CalculaDigitoCNPJ(msk_cgc.Text) Then
            msk_cgc.SetFocus
        End If
    End If
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_nome.SetFocus
    End If
End Sub
Private Sub txt_cidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data.SetFocus
    End If
End Sub
Private Sub txt_cpf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_identidade.SetFocus
    End If
End Sub
Private Sub txt_identidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_orgao_emissor.SetFocus
    End If
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub txt_orgao_emissor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_cidade.SetFocus
    End If
End Sub
Private Sub txt_proveniente_GotFocus()
    txt_proveniente.SelStart = 0
    txt_proveniente.SelLength = Len(txt_proveniente)
End Sub
Private Sub txt_proveniente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor.SetFocus
    End If
End Sub
Private Sub txt_razao_social_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_cgc.SetFocus
    End If
End Sub
Private Sub txt_valor_GotFocus()
    txt_valor.SelStart = 0
    txt_valor.SelLength = Len(txt_valor)
End Sub
Private Sub txt_valor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_cpf.SetFocus
    End If
End Sub
Private Sub txt_valor_LostFocus()
    txt_valor = Format(txt_valor, "###,##0.00")
End Sub
