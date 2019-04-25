VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form emissao_cheque_formulario 
   Caption         =   "Emissão de Cheques (formulário)"
   ClientHeight    =   3735
   ClientLeft      =   1200
   ClientTop       =   1590
   ClientWidth     =   6975
   Icon            =   "emi_cheque_formulario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3735
   ScaleWidth      =   6975
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2820
      Picture         =   "emi_cheque_formulario.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2760
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "emi_cheque_formulario.frx":199C
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Exclui o registro atual."
      Top             =   2760
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "emi_cheque_formulario.frx":302E
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Altera o registro atual."
      Top             =   2760
      Width           =   795
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "emi_cheque_formulario.frx":4528
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Cria um novo registro."
      Top             =   2760
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3720
      Picture         =   "emi_cheque_formulario.frx":5BBA
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Imprime cheque em formulário."
      Top             =   2760
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "emi_cheque_formulario.frx":71C4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt_nome 
         Height          =   300
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   12
         Top             =   1740
         Width           =   4935
      End
      Begin VB.Data dta_funcionario 
         Caption         =   "dta_funcionario"
         Connect         =   "Access"
         DatabaseName    =   "Sgp_data.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3900
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Funcionario"
         Top             =   1380
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.TextBox txt_historico 
         Height          =   300
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   14
         Top             =   2100
         Width           =   4935
      End
      Begin VB.TextBox txt_numero 
         Height          =   300
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox msk_valor 
         Height          =   300
         Left            =   1680
         TabIndex        =   7
         Top             =   1020
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDBCtls.DBCombo dbcbo_funcionario 
         Bindings        =   "emi_cheque_formulario.frx":849E
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   1740
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   ""
      End
      Begin VB.Label lbl_extenso 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1680
         TabIndex        =   9
         Top             =   1380
         Width           =   4935
      End
      Begin VB.Label Label2 
         Caption         =   "&Histórico"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2100
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Extenso"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1380
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "&Favorecido"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "&Número do Cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &Emissão"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "&Valor do Cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1020
         Width           =   1455
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4680
      TabIndex        =   22
      Top             =   2640
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "emi_cheque_formulario.frx":84BC
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "emi_cheque_formulario.frx":99B6
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "emi_cheque_formulario.frx":AEB0
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "emi_cheque_formulario.frx":C322
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6060
      Picture         =   "emi_cheque_formulario.frx":D8A4
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cancela o registro atual."
      Top             =   2760
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5160
      Picture         =   "emi_cheque_formulario.frx":ED9E
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Confirma o registro atual."
      Top             =   2760
      Width           =   795
   End
End
Attribute VB_Name = "emissao_cheque_formulario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_emissao_cheque_formulario As Integer
Dim lOpcao As Integer
Dim l_data As Date
Dim l_numero As String
Dim l_old_historico As String
Dim tbl_cheque_formulario As Table
Dim tbl_funcionario As Table
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
'    cmd_pesquisa.Enabled = True
    cmd_excluir.Enabled = True
    cmd_sair.Enabled = True
    cmd_imprimir.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
End Sub
Private Sub AtualTabe()
    l_data = msk_data.Text
    l_numero = Val(txt_numero.Text)
    With tbl_cheque_formulario
        !Empresa = g_empresa
        ![Numero do Cheque] = txt_numero.Text
        ![Data de Emissao] = msk_data.Text
        !valor = fValidaValor2(msk_valor.Text)
        ![Nome do Favorecido] = txt_nome.Text
        !Historico = txt_historico.Text
    End With
End Sub
Private Sub AtualTela()
    With tbl_cheque_formulario
        l_data = ![Data de Emissao]
        l_numero = ![Numero do Cheque]
        txt_numero.Text = ![Numero do Cheque]
        msk_data.Text = Format(![Data de Emissao], "dd/mm/yyyy")
        msk_valor.Text = Format(!valor, "###,##0.00")
        lbl_extenso.Caption = FazExtenso(fValidaValor2(msk_valor))
        txt_nome.Text = ![Nome do Favorecido]
        txt_historico.Text = Format(!Historico, "")
    End With
    frm_dados.Enabled = False
End Sub
Function BuscaDados() As Boolean
    BuscaDados = False
    If tbl_cheque_formulario.RecordCount > 0 Then
        tbl_cheque_formulario.Seek "<", g_empresa, CDate("31/12/2500"), "999999"
        If Not tbl_cheque_formulario.NoMatch Then
            If tbl_cheque_formulario!Empresa = g_empresa Then
                AtualTela
                BuscaDados = True
            End If
        End If
    End If
End Function
Function BuscaRegistro(x_data As Date, x_numero As String) As Boolean
    BuscaRegistro = False
    tbl_cheque_formulario.Seek "=", g_empresa, x_data, x_numero
    If Not tbl_cheque_formulario.NoMatch Then
        BuscaRegistro = True
        AtualTela
    End If
End Function
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_excluir.Enabled = False
'    cmd_pesquisa.Enabled = False
    cmd_sair.Enabled = False
    cmd_imprimir.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_cheque_formulario.Close
    tbl_funcionario.Close
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    CriaNovoNumero
End Sub
Private Sub CriaNovoNumero()
    txt_numero = 0
    If tbl_cheque_formulario.RecordCount > 0 Then
        tbl_cheque_formulario.Seek "<", g_empresa, CDate("31/12/2500"), "999999"
        If Not tbl_cheque_formulario.NoMatch Then
            If tbl_cheque_formulario!Empresa = g_empresa Then
                txt_numero = tbl_cheque_formulario![Numero do Cheque]
            End If
        End If
    End If
    txt_numero = CLng(txt_numero) + 1
End Sub
Private Sub Relatorio()
    Dim posicao_y As Currency
    Dim tamanho_form As Integer
    Dim largura_form As Integer
    Dim x_tamanho_string As Currency
    Dim x_extenso As String
    Dim x_asterisco As String
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    'Seleciona Formulário de cheque
    Printer.PaperSize = 256
    'Seleciona largura do formulário
'    Printer.ScaleWidth = 20
    largura_form = Printer.ScaleWidth
    'Seleciona altura do formulário
'    Printer.ScaleHeight = 8
    tamanho_form = Printer.ScaleHeight
    'Seleciona nome da fonte
    Printer.FontName = "Arial"
    Printer.FontName = "Roman"
    
    Printer.CurrentY = 0.9
    Printer.CurrentX = 2
    Printer.Print "." 'x_extenso
    
    Printer.CurrentY = 0
    Printer.FontSize = 16
    Printer.CurrentX = 14
    Printer.Print msk_valor
    Printer.Print

    'Define Tamanho da fonte
    Printer.FontSize = 14
    x_tamanho_string = Printer.TextWidth(lbl_extenso.Caption)
    'Verifica se o tamanho ultrapassa a 14.3 centimetros
    If x_tamanho_string > 14.3 Then
        'Define Tamanho da fonte menor
        Printer.FontSize = 12
        x_tamanho_string = Printer.TextWidth(lbl_extenso.Caption)
        'Verifica se o tamanho ultrapassa a 14.3 centimetros
        If x_tamanho_string > 14.3 Then
            MsgBox "Tamanho da String = " & x_tamanho_string, 64, "Extenso muito longo!"
            Printer.EndDoc
            Exit Sub
        End If
    End If
    
    x_extenso = "( " & lbl_extenso
    Do Until x_tamanho_string >= 14.3
        x_extenso = x_extenso & "*"
        x_tamanho_string = Printer.TextWidth(x_extenso)
    Loop
    
    Printer.CurrentY = 0.9
    Printer.CurrentX = 2
    Printer.Print x_extenso

    Printer.FontSize = 14
    Printer.CurrentY = 1.8
    Printer.CurrentX = 0
    x_tamanho_string = 0
    x_asterisco = ""
    Do Until x_tamanho_string >= 16.2
        x_asterisco = x_asterisco & "*"
        x_tamanho_string = Printer.TextWidth(x_asterisco)
    Loop
    Printer.Print x_asterisco
    
    Printer.CurrentY = 2.5
    Printer.CurrentX = 0.5
    Printer.Print txt_nome
        
    Printer.CurrentY = 3.4
    Printer.CurrentX = 8.5
    Printer.Print "Goiânia"
    Printer.CurrentY = 3.4
    Printer.CurrentX = 10.8
    Printer.Print Day(msk_data.Text)
    Printer.CurrentY = 3.4
    Printer.CurrentX = 12
    Printer.Print Format(CDate(msk_data.Text), "mmmm")
    Printer.CurrentY = 3.4
    Printer.CurrentX = 15.8
    Printer.Print Mid(Year(CDate(msk_data.Text)), 3, 2)
    Printer.EndDoc
End Sub
Private Sub SelecionaImpressora()
Dim Impressora As Printer
    For Each Impressora In Printers
        If Impressora.DeviceName = "Epson FX-1170" Or Impressora.DeviceName = "\\Pc-analu\EPSON FX1170" Then
            Set Printer = Impressora
        End If
    Next
End Sub
Private Sub cmd_alterar_Click()
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frm_dados.Enabled = True
    txt_nome.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If tbl_cheque_formulario.RecordCount > 0 Then
        tbl_cheque_formulario.MovePrevious
        If Not tbl_cheque_formulario.BOF Then
            If tbl_cheque_formulario!Empresa = g_empresa Then
                AtualTela
                Exit Sub
            End If
        End If
        MsgBox "Início de Arquivo.", 48, "Atenção!"
        tbl_cheque_formulario.MoveNext
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    tbl_cheque_formulario.Seek ">", g_empresa, ""
    If Not tbl_cheque_formulario.NoMatch Then
        AtivaBotoes
        BuscaDados
        cmd_alterar.SetFocus
    Else
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
        LimpaTela
        cmd_novo.SetFocus
    End If
End Sub
Private Sub LimpaTela()
    txt_numero.Text = ""
    msk_data.Text = "__/__/____"
    msk_valor.Text = ""
    lbl_extenso.Caption = ""
    txt_nome.Text = ""
    txt_historico.Text = ""
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data.Text = RetiraGString(1)
    msk_valor.SetFocus
    g_string = " "
End Sub
Private Sub cmd_excluir_Click()
    If tbl_cheque_formulario![Numero do Cheque] <> "" Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = vbYes Then
            tbl_cheque_formulario.Edit
            tbl_cheque_formulario.Delete
            LimpaTela
            If Not BuscaDados Then
                DesativaBotoes
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
                cmd_novo.SetFocus
            End If
        End If
    End If
End Sub
Private Sub cmd_imprimir_Click()
    If SelecionaImpressoraHP(Me) Then
        lbl_extenso = FazExtenso(fValidaValor2(msk_valor.Text))
        Call GravaAuditoria(1, Me.name, 7, "")
        Relatorio
        cmd_novo.SetFocus
    End If
End Sub
Private Sub cmd_novo_Click()
    LimpaTela
    Inclui
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
    End If
    frm_dados.Enabled = True
    msk_valor.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo ErrorFile
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            tbl_cheque_formulario.AddNew
            AtualTabe
            tbl_cheque_formulario.Update
            cmd_imprimir.SetFocus
        ElseIf lOpcao = 2 Then
            tbl_cheque_formulario.Edit
            AtualTabe
            tbl_cheque_formulario.Update
            cmd_imprimir.SetFocus
        End If
        Call BuscaRegistro(l_data, l_numero)
        AtualTela
        cmd_imprimir.SetFocus
    End If
    Exit Sub
ErrorFile:
    ErroArquivo tbl_cheque_formulario.name, "Cheque Formulárioo"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not txt_numero.Text <> "" Then
        MsgBox "Informe o número do cheque.", 64, "Atenção!"
        txt_numero.SetFocus
    ElseIf Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data de emissão.", 64, "Atenção!"
        msk_data.SetFocus
    ElseIf Not fValidaValor2(msk_valor.Text) > 0 Then
        MsgBox "Informe o valor do cheque.", 64, "Atenção!"
        msk_valor.SetFocus
    ElseIf Not txt_nome.Text <> "" Then
        MsgBox "Informe o favorecido.", 64, "Atenção!"
        txt_nome.SetFocus
'    ElseIf Not txt_historico.Text <> "" Then
'        MsgBox "Informe o histórico.", 64, "Atenção!"
'        txt_historico.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_primeiro_Click()
    If tbl_cheque_formulario.RecordCount Then
        tbl_cheque_formulario.Seek ">", g_empresa, CDate("01/01/1900"), ""
        If Not tbl_cheque_formulario.NoMatch Then
            If tbl_cheque_formulario!Empresa = g_empresa Then
                AtualTela
                cmd_proximo.SetFocus
                Exit Sub
            End If
        End If
        MsgBox "Não há registros nesta empresa.", 64, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If tbl_cheque_formulario.RecordCount > 0 Then
        tbl_cheque_formulario.MoveNext
        If Not tbl_cheque_formulario.EOF Then
            If tbl_cheque_formulario!Empresa = g_empresa Then
                AtualTela
                Exit Sub
            End If
        End If
        MsgBox "Início de Arquivo.", 48, "Atenção!"
        tbl_cheque_formulario.MovePrevious
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If tbl_cheque_formulario.RecordCount Then
        tbl_cheque_formulario.Seek "<", g_empresa, CDate("31/12/2500"), "999999"
        If Not tbl_cheque_formulario.NoMatch Then
            If tbl_cheque_formulario!Empresa = g_empresa Then
                AtualTela
                cmd_anterior.SetFocus
                Exit Sub
            End If
        End If
        MsgBox "Não há registros nesta empresa.", 64, "Erro de Verificação!"
    End If
End Sub
Private Sub dbcbo_funcionario_GotFocus()
    txt_nome.Visible = False
    dbcbo_funcionario.BoundText = ""
End Sub
Private Sub dbcbo_funcionario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        txt_historico.SetFocus
    End If
End Sub
Private Sub dbcbo_funcionario_LostFocus()
    txt_nome.Visible = True
    If dbcbo_funcionario.BoundText <> "" Then
        txt_nome.Text = dbcbo_funcionario
        txt_historico.SetFocus
    Else
        txt_nome.Text = ""
        txt_nome.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If flag_emissao_cheque_formulario = 0 Then
        DesativaBotoes
        If BuscaDados Then
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        cmd_novo.SetFocus
    Else
        flag_emissao_cheque_formulario = 0
    End If
    dta_funcionario.RecordSource = "SELECT Codigo, Nome FROM Funcionario WHERE Empresa = " & g_empresa & " AND Situacao = 'A' ORDER BY Nome"
    dta_funcionario.Refresh
    Screen.MousePointer = 1
End Sub
Private Sub Form_Deactivate()
    flag_emissao_cheque_formulario = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And lOpcao = 0 Then
        KeyCode = 0
        cmd_novo_Click
    ElseIf KeyCode = vbKeyF3 And lOpcao = 0 Then
        KeyCode = 0
        cmd_alterar_Click
    ElseIf KeyCode = vbKeyF4 And Shift = 0 And lOpcao = 0 Then
        KeyCode = 0
        cmd_excluir_Click
    ElseIf KeyCode = vbKeyF6 And lOpcao = 0 Then
        KeyCode = 0
        cmd_imprimir_Click
    ElseIf KeyCode = vbKeyF7 And lOpcao = 0 Then
        KeyCode = 0
        cmd_primeiro_Click
    ElseIf KeyCode = vbKeyF8 And lOpcao = 0 Then
        KeyCode = 0
        cmd_anterior_Click
    ElseIf KeyCode = vbKeyF9 And lOpcao = 0 Then
        KeyCode = 0
        cmd_proximo_Click
    ElseIf KeyCode = vbKeyF10 And lOpcao = 0 Then
        KeyCode = 0
        cmd_ultimo_Click
    ElseIf KeyCode = vbKeyF11 And lOpcao > 0 Then
        KeyCode = 0
        cmd_ok_Click
    ElseIf KeyCode = vbKeyF12 And lOpcao > 0 Then
        KeyCode = 0
        cmd_cancelar_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    Set tbl_cheque_formulario = bd_sgp.OpenTable("Cheque_Formulario")
    Set tbl_funcionario = bd_sgp.OpenTable("Funcionario")
    tbl_cheque_formulario.Index = "id_data"
    SelecionaImpressora
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_GotFocus()
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
    End If
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_valor.SetFocus
    End If
End Sub
Private Sub msk_data_LostFocus()
    If IsDate(msk_data.Text) Then
        g_data_def = msk_data.Text
    End If
End Sub
Private Sub msk_valor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        dbcbo_funcionario.SetFocus
    End If
End Sub
Private Sub msk_valor_LostFocus()
    If Val(msk_valor.Text) > 0 Then
        msk_valor.Text = Format(msk_valor.Text, "###,##0.00")
        lbl_extenso.Caption = FazExtenso(fValidaValor2(msk_valor.Text))
    End If
End Sub
Private Sub txt_historico_GotFocus()
    If lOpcao = 1 Then
        txt_historico.Text = l_old_historico
        txt_historico.SelLength = Len(txt_historico.Text)
    End If
End Sub
Private Sub txt_historico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_historico_LostFocus()
    l_old_historico = txt_historico.Text
End Sub
Private Sub txt_nome_Click()
    dbcbo_funcionario.SetFocus
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_historico.SetFocus
    End If
End Sub
Private Sub txt_numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
