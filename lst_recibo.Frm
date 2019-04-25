VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form emissao_recibo 
   Caption         =   "Emissão de Recibo"
   ClientHeight    =   3855
   ClientLeft      =   1965
   ClientTop       =   3705
   ClientWidth     =   6975
   Icon            =   "lst_recibo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_recibo.frx":030A
   ScaleHeight     =   3855
   ScaleWidth      =   6975
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   1200
      Picture         =   "lst_recibo.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Cria um Novo Recibo."
      Top             =   2880
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3060
      Picture         =   "lst_recibo.frx":19E2
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Imprime o Recibo."
      Top             =   2880
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4920
      Picture         =   "lst_recibo.frx":2FEC
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2880
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2715
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6735
      Begin VB.TextBox txt_empresa 
         Height          =   300
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   3
         Top             =   540
         Width           =   4875
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_recibo.frx":467E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txt_funcionario 
         Height          =   300
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   11
         Top             =   1800
         Width           =   4875
      End
      Begin VB.CheckBox chk_formulario 
         Caption         =   "&Recibo em Formulário"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox txt_valor 
         Height          =   300
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1380
         Width           =   1095
      End
      Begin VB.TextBox txt_proveniente 
         Height          =   300
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   13
         Top             =   2220
         Width           =   4875
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   300
         Left            =   1680
         TabIndex        =   5
         Top             =   960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc adodcFuncionario 
         Height          =   330
         Left            =   3480
         Top             =   1200
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adodcFuncionario"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo dtcboFuncionario 
         Bindings        =   "lst_recibo.frx":5958
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   1800
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboFuncionario"
      End
      Begin VB.Label Label4 
         Caption         =   "Emp&resa"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "&Valor"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1380
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "&Proveniente a"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   2220
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "N&ome"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   180
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_recibo"
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
    ImprimeTexto txt_empresa.Text, lColunaI + 6.5, lColuna(1), lLinhaTab, lLocal
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
    Printer.Line (lColuna(0), lLinhaI)-(lColuna(1), lLinhaI)
    Printer.Line (lColuna(0), lLinhaI)-(lColuna(0), lLinhaI + 10)
    Printer.Line (lColuna(1), lLinhaI)-(lColuna(1), lLinhaI + 10)
    Printer.Line (lColuna(0), lLinhaI + 10)-(lColuna(1), lLinhaI + 10)
    Printer.FontSize = 20
    Printer.DrawWidth = 8
    Printer.FontBold = True
    ImprimeCentralizado "R E C I B O", lColuna(0), lColuna(1), lLinhaI + 1, lLocal
    Printer.DrawWidth = 1
    Printer.Line (lColuna(0), lLinhaI + 11.7)-(lColuna(1), lLinhaI + 11.7)
End Sub
Private Sub PosicionaProximoRegistro()
    adodcFuncionario.Recordset.Find ("Codigo = " & Val(Me.dtcboFuncionario.BoundText))
    If Not adodcFuncionario.Recordset.EOF Then
        adodcFuncionario.Recordset.MoveNext
        If Not adodcFuncionario.Recordset.EOF Then
            dtcboFuncionario.BoundText = adodcFuncionario.Recordset!Codigo
        Else
            dtcboFuncionario.BoundText = ""
        End If
    End If
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    If chk_formulario Then
        ImprimeDadosFormulario
    Else
        ImpDados
    End If
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
    ImprimeTexto txt_empresa.Text, 4, 18, 3.7, 1
    
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
    txt_valor.SetFocus
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
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf Not fValidaValor2(txt_valor.Text) > 0 Then
        MsgBox "Informe o valor.", vbInformation, "Atenção!"
        txt_valor.SetFocus
    ElseIf Not txt_funcionario.Text > "" Then
        MsgBox "Escolha o funcionario.", vbInformation, "Atenção!"
        dtcboFuncionario.SetFocus
    ElseIf Not txt_proveniente.Text > "" Then
        MsgBox "Informe a referência.", vbInformation, "Atenção!"
        txt_proveniente.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_novo_Click()
    txt_funcionario = ""
    PosicionaProximoRegistro
    txt_valor.SetFocus
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub dtcboFuncionario_GotFocus()
    txt_funcionario.Visible = False
End Sub
Private Sub dtcboFuncionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_proveniente.SetFocus
    End If
End Sub
Private Sub dtcboFuncionario_LostFocus()
    txt_funcionario.Visible = True
    If dtcboFuncionario.BoundText <> "" Then
        txt_funcionario = dtcboFuncionario
        txt_proveniente.SetFocus
    Else
        txt_funcionario = ""
        txt_funcionario.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    TabelaFuncionarioRefresh
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        txt_valor.SetFocus
    End If
    Screen.MousePointer = 1
End Sub
Private Sub TabelaFuncionarioRefresh()
'    dta_funcionario.RecordSource = "SELECT Codigo, Nome FROM Funcionario WHERE Empresa = " & g_empresa & " AND Situacao = " & preparaTexto("A") & " ORDER BY Nome"
'    dta_funcionario.Refresh
    Set adodcFuncionario.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Funcionario WHERE Empresa = " & g_empresa & " AND Situacao = " & preparaTexto("A") & " ORDER BY Nome")
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        KeyCode = 0
        cmd_novo_Click
    ElseIf KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_imprimir_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    txt_empresa.Text = g_nome_empresa
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
Private Sub txt_empresa_GotFocus()
    txt_empresa.SelStart = 0
    txt_empresa.SelLength = Len(txt_empresa.Text)
End Sub
Private Sub txt_empresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data.SetFocus
    End If
End Sub
Private Sub txt_funcionario_Click()
    dtcboFuncionario.SetFocus
End Sub
Private Sub txt_funcionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_proveniente.SetFocus
    End If
End Sub
Private Sub txt_proveniente_GotFocus()
    txt_proveniente.SelStart = 0
    txt_proveniente.SelLength = Len(txt_proveniente.Text)
End Sub
Private Sub txt_proveniente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
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
        dtcboFuncionario.SetFocus
    End If
End Sub
Private Sub txt_valor_LostFocus()
    txt_valor = Format(txt_valor, "###,##0.00")
End Sub
