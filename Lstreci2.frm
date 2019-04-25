VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_emissao_recibo_folhas 
   Caption         =   "Emissão dos Recibos (Folhas)"
   ClientHeight    =   3045
   ClientLeft      =   1815
   ClientTop       =   3285
   ClientWidth     =   3750
   Icon            =   "Lstreci2.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Lstreci2.frx":030A
   ScaleHeight     =   3045
   ScaleWidth      =   3750
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2220
      Picture         =   "Lstreci2.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2040
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   720
      Picture         =   "Lstreci2.frx":19E2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprime o Recibo de Cheque (Folha)."
      Top             =   2040
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "Lstreci2.frx":2FEC
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   540
         Width           =   495
      End
      Begin VB.CheckBox chk_formulario 
         Caption         =   "&Recibo em Formulário"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox txt_numero_f 
         Height          =   285
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   8
         Top             =   1380
         Width           =   735
      End
      Begin VB.TextBox txt_numero_i 
         Height          =   285
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   540
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
         Caption         =   "Cheque &Final"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1380
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "&Cheque Inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "&Data de Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Width           =   1515
      End
   End
   Begin Crystal.CrystalReport cr_relato 
      Left            =   120
      Top             =   1980
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\VB5\Sgp\Data\lst_baixa_contas_a_pagar_fornecedor.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frm_emissao_recibo_folhas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_linha_tab As Integer
Dim l_coluna_tab As Integer
Dim l_linha As Integer
Dim lSQL As String

Dim rsChequeFolha As New ADODB.Recordset

Private Sub ZeraVariaveis()
    l_coluna_tab = 0
    l_linha_tab = 0
    l_linha = 0
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set rsChequeFolha = Nothing
End Sub
Private Sub ImprimeDados()
    Dim x_extenso As String
    Dim x_mes As String
    Dim x_data As Date
    Dim x2_data As String
    Dim x_tamanho_string As Currency
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    Printer.FontName = "Arial"
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    Do Until rsChequeFolha.EOF
'        If l_linha = 2 Then
'            ImprimeGrade l_linha
'            Printer.NewPage
'            l_linha = 0
'        End If
        'teste
        Printer.PSet (0, 0)
        l_linha = l_linha + 1
        l_linha_tab = l_linha * 10 - 10 + 2 + l_linha_tab
        If l_linha = 2 Then
            l_linha_tab = l_linha_tab + 1.9
        End If
        l_coluna_tab = 0
        Printer.FontSize = 14
        Printer.DrawWidth = 6
        Printer.CurrentX = l_coluna_tab + 15
        Printer.CurrentY = l_linha_tab + 2
        Printer.Print "R$ " & Format(rsChequeFolha!valor, "###,##0.00")
        Printer.FontSize = 10
        Printer.DrawWidth = 4
        Printer.CurrentX = l_coluna_tab + 3
        Printer.CurrentY = l_linha_tab + 3
        Printer.Print "Recebi da empresa, "
        Printer.FontSize = 14
        Printer.DrawWidth = 6
        Printer.CurrentX = l_coluna_tab + 6.3
        Printer.CurrentY = l_linha_tab + 2.9
        Printer.Print g_nome_empresa
        Printer.FontSize = 10
        Printer.DrawWidth = 4
        Printer.CurrentX = l_coluna_tab + 1
        Printer.CurrentY = l_linha_tab + 4
        Printer.Print "A quantia supra de "
        x_extenso = FazExtenso(rsChequeFolha!valor)
        Printer.CurrentX = l_coluna_tab + 4.1
        Printer.CurrentY = l_linha_tab + 4
        Printer.Print x_extenso
        Printer.CurrentX = l_coluna_tab + 1
        Printer.CurrentY = l_linha_tab + 5
        Printer.Print "Proveniente de "
        Printer.CurrentX = l_coluna_tab + 3.55
        Printer.CurrentY = l_linha_tab + 5
        Printer.Print rsChequeFolha!Historico
        x_data = msk_data
        x2_data = msk_data
        FazExtensoMes x2_data, x_mes
        Printer.CurrentX = l_coluna_tab + 6
        Printer.CurrentY = l_linha_tab + 6
        Printer.Print g_cidade_empresa & ","
        Printer.CurrentX = l_coluna_tab + 7.5
        Printer.CurrentY = l_linha_tab + 6
        Printer.Print Day(msk_data)
        Printer.CurrentX = l_coluna_tab + 8.5
        Printer.CurrentY = l_linha_tab + 6
        Printer.Print "de"
        Printer.CurrentX = l_coluna_tab + 9.5
        Printer.CurrentY = l_linha_tab + 6
        Printer.Print x_mes
        Printer.CurrentX = l_coluna_tab + 12
        Printer.CurrentY = l_linha_tab + 6
        Printer.Print "de"
        Printer.CurrentX = l_coluna_tab + 13
        Printer.CurrentY = l_linha_tab + 6
        Printer.Print Year(msk_data)
        Printer.CurrentX = l_coluna_tab + 5
        Printer.CurrentY = l_linha_tab + 8.6
        Printer.Print "___________________________________________________"
        x_tamanho_string = Printer.TextWidth(rsChequeFolha!Nome)
        Printer.CurrentX = l_coluna_tab + ((20 - x_tamanho_string) / 2)
        Printer.CurrentY = l_linha_tab + 9
        Printer.Print rsChequeFolha!Nome
'
        ImprimeGrade l_linha
        Printer.NewPage
        l_linha = 0
'
        rsChequeFolha.MoveNext
    Loop
'    ImprimeGrade l_linha
    Printer.EndDoc
End Sub
Private Sub ImprimeDadosFormulario()
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    Printer.FontName = "Arial"
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    Do Until rsChequeFolha.EOF
        If l_linha > 0 Then
            Printer.EndDoc
        End If
        l_linha = l_linha + 1
        
        Printer.FontSize = 14
        Printer.DrawWidth = 6
        ImprimeCentralizado Format(rsChequeFolha!valor, "###,##0.00"), 12.1, 18.1, 2.7, 1
        
        Printer.FontSize = 14
        Printer.DrawWidth = 6
        ImprimeTexto g_nome_empresa, 4, 18, 4#, 1
        
        Printer.FontSize = 10
        Printer.DrawWidth = 4
        ImprimeTexto FazExtenso(rsChequeFolha!valor), 4, 18, 5.2, 1
        
        ImprimeTexto rsChequeFolha!Historico, 4, 18, 7.4, 1
        
        ImprimeCentralizado Trim(g_cidade_empresa) & ", " & Day(msk_data) & " de " & Format(msk_data, "mmmm") & " de " & Format(msk_data, "yyyy") & ".", 8, 18, 11.4, 1
        
        ImprimeCentralizado rsChequeFolha!Nome, 8, 18, 12.6, 1
        rsChequeFolha.MoveNext
    Loop
    Printer.EndDoc
End Sub
Private Sub ImprimeGrade(x_vezes As Integer)
    Dim i As Integer
    Dim x_tamanho_string As Currency
    'Seleciona tamanho da fonte
    l_coluna_tab = 0
    l_linha_tab = 0
    Printer.FontSize = 20
    Printer.DrawWidth = 8
    For i = 1 To l_linha
        l_linha_tab = i * 10 - 10 + 2 + l_linha_tab
        x_tamanho_string = Printer.TextWidth("R E C I B O")
        Printer.CurrentX = l_coluna_tab + ((20 - x_tamanho_string) / 2)
        Printer.CurrentY = l_linha_tab + 0.8
        Printer.Print "R E C I B O"
        Printer.Line (l_coluna_tab, l_linha_tab)-(l_coluna_tab + 20, l_linha_tab)
        Printer.Line (l_coluna_tab, l_linha_tab + 10)-(l_coluna_tab + 20, l_linha_tab + 10)
        Printer.Line (l_coluna_tab, l_linha_tab)-(l_coluna_tab, l_linha_tab + 10)
        Printer.Line (l_coluna_tab + 20, l_linha_tab)-(l_coluna_tab + 20, l_linha_tab + 10)
        If i = 1 Then
            Printer.DrawWidth = 1
            Printer.Line (l_coluna_tab, l_linha_tab + 10 + 1.7)-(l_coluna_tab + 20, l_linha_tab + 10 + 1.7)
            Printer.DrawWidth = 8
        End If
        l_linha_tab = l_linha_tab + 1.9
    Next
    l_coluna_tab = 0
    l_linha_tab = 0
    Printer.FontSize = 10
    Printer.DrawWidth = 4
End Sub
Private Sub Relatorio()
    Dim i As Integer
    ZeraVariaveis
'    tbl_cheque_folha.Index = "id_numero"
'    tbl_cheque_folha.Seek ">", g_empresa, Format((txt_numero_i - 1), "000000")
'    If Not tbl_cheque_folha.NoMatch Then
'        If tbl_cheque_folha!Empresa = g_empresa Then
'            If chk_formulario Then
'                ImprimeDadosFormulario
'            Else
'                ImprimeDados
'            End If
'        End If
'    End If
    lSQL = ""
    lSQL = lSQL & " SELECT Valor, Historico, Nome"
    lSQL = lSQL & "   FROM Cheque_Folha"
    lSQL = lSQL & "  WHERE Empresa = " & g_empresa
    lSQL = lSQL & "    AND Numero >= " & preparaTexto(txt_numero_i.Text)
    lSQL = lSQL & "    AND Numero <= " & preparaTexto(txt_numero_f.Text)
    
    Set rsChequeFolha = Conectar.RsConexao(lSQL)
    If rsChequeFolha.RecordCount > 0 Then
        If chk_formulario Then
            ImprimeDadosFormulario
        Else
            ImprimeDados
        End If
    End If
    rsChequeFolha.Close
    Set rsChequeFolha = Nothing
    
    cmd_sair.SetFocus
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
    txt_numero_i.SetFocus
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    Dim x_flag As Integer
    If ValidaCampos Then
        If SelecionaImpressoraHP(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        End If
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data de emissão.", 64, "Atenção!"
        msk_data.SetFocus
    ElseIf Not Val(txt_numero_i) > 0 Then
        MsgBox "Informe o número inicial.", 64, "Atenção!"
        txt_numero_i.SetFocus
    ElseIf Not Val(txt_numero_f) > 0 Then
        MsgBox "Informe o número final.", 64, "Atenção!"
        txt_numero_f.SetFocus
    ElseIf txt_numero_f < txt_numero_i Then
        MsgBox "Número final deve ser maior que o número inicial.", 64, "Atenção!"
        txt_numero_f.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Finaliza
    Unload Me
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        txt_numero_i.SetFocus
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_imprimir_Click
    End If
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
End Sub
Private Sub msk_data_GotFocus()
    msk_data.SelStart = 0
    msk_data.SelLength = 5
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_numero_i.SetFocus
    End If
End Sub
Private Sub txt_numero_f_GotFocus()
    txt_numero_f.SelStart = 0
    txt_numero_f.SelLength = Len(txt_numero_f)
End Sub
Private Sub txt_numero_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_numero_i_GotFocus()
    txt_numero_i.SelStart = 0
    txt_numero_i.SelLength = Len(txt_numero_i)
End Sub
Private Sub txt_numero_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_numero_f.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
