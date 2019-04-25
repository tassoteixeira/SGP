VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form config_liberacao_caixa 
   Caption         =   "Liberação da Digitação"
   ClientHeight    =   2460
   ClientLeft      =   1125
   ClientTop       =   1350
   ClientWidth     =   7035
   Icon            =   "config_liberacao_caixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "config_liberacao_caixa.frx":030A
   ScaleHeight     =   2460
   ScaleWidth      =   7035
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5040
      Picture         =   "config_liberacao_caixa.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1500
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   3120
      Picture         =   "config_liberacao_caixa.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cancela o registro atual."
      Top             =   1500
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   1200
      Picture         =   "config_liberacao_caixa.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Confirma o registro atual."
      Top             =   1500
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6795
      Begin VB.ComboBox cboTipoMovimento 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   180
         Width           =   2175
      End
      Begin VB.ComboBox cbo_periodo_f 
         Height          =   315
         Left            =   5580
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   900
         Width           =   495
      End
      Begin VB.ComboBox cbo_periodo_i 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   900
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   5580
         TabIndex        =   6
         Top             =   540
         Width           =   1095
         _ExtentX        =   1931
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
         TabIndex        =   4
         Top             =   540
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "&Tipo do Movimento"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Periodo &final"
         Height          =   255
         Index           =   3
         Left            =   4020
         TabIndex        =   9
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Periodo inicial"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   4020
         TabIndex        =   5
         Top             =   600
         Width           =   1515
      End
   End
End
Attribute VB_Name = "config_liberacao_caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Configuracao As New cConfiguracao
Private LiberacaoDigitacao As New cLiberacaoDigitacao
Private ConfiguracaoDiversa As New cConfiguracaoDiversa

Private Sub AtualTabe()
    LiberacaoDigitacao.Empresa = g_empresa
    LiberacaoDigitacao.DataInicial = msk_data_i.Text
    LiberacaoDigitacao.DataFinal = msk_data_f.Text
    LiberacaoDigitacao.PeriodoInicial = cbo_periodo_i.Text
    LiberacaoDigitacao.PeriodoFinal = cbo_periodo_f.Text
    LiberacaoDigitacao.TipoMovimento = Val(cboTipoMovimento.Text)
End Sub
Private Sub AtualTela()
    msk_data_i.Text = Format(LiberacaoDigitacao.DataInicial, "dd/mm/yyyy")
    msk_data_f.Text = Format(LiberacaoDigitacao.DataFinal, "dd/mm/yyyy")
    cbo_periodo_i.ListIndex = LiberacaoDigitacao.PeriodoInicial - 1
    cbo_periodo_f.ListIndex = LiberacaoDigitacao.PeriodoInicial - 1
    If Configuracao.LocalizarCodigo(g_empresa) Then
        If cbo_periodo_f.ListCount <= LiberacaoDigitacao.PeriodoFinal - 1 Then
            cbo_periodo_f.ListIndex = Configuracao.QuantidadePeriodos - 1
        Else
            cbo_periodo_f.ListIndex = LiberacaoDigitacao.PeriodoFinal - 1
        End If
    End If
End Sub
Private Sub CriaNovoRegistro()
    LiberacaoDigitacao.Empresa = g_empresa
    LiberacaoDigitacao.DataInicial = g_data_def
    LiberacaoDigitacao.DataFinal = g_data_def
    LiberacaoDigitacao.PeriodoInicial = 1
    LiberacaoDigitacao.PeriodoFinal = 1
    LiberacaoDigitacao.TipoMovimento = Val(cboTipoMovimento.Text)
    If LiberacaoDigitacao.Incluir Then
        AtualTela
    Else
        MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Configuracao = Nothing
    Set LiberacaoDigitacao = Nothing
End Sub
Private Sub cbo_periodo_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub cbo_periodo_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_f.SetFocus
    End If
End Sub
Private Sub cbo_periodo_i_LostFocus()
    If cbo_periodo_i.ListIndex <> -1 Then
        cbo_periodo_f.ListIndex = cbo_periodo_i.ListIndex
    End If
End Sub
Private Sub cboTipoMovimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_i.SetFocus
    End If
End Sub
Private Sub cboTipoMovimento_LostFocus()
    If cboTipoMovimento.ListIndex <> -1 Then
        If LiberacaoDigitacao.LocalizarCodigo(g_empresa, Val(cboTipoMovimento.Text)) Then
            AtualTela
        Else
            CriaNovoRegistro
        End If
    End If
End Sub
Private Sub cmd_cancelar_Click()
    Call GravaAuditoria(1, Me.name, 9, "")
    If LiberacaoDigitacao.LocalizarCodigo(g_empresa, Val(cboTipoMovimento.Text)) Then
        AtualTela
    Else
        MsgBox "Registro inexistente!", vbInformation, "Erro de Verificação!"
    End If
    cmd_sair.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        Call GravaAuditoria(1, Me.name, 10, "De: Dt I:" & fMascaraData(LiberacaoDigitacao.DataInicial) & " - " & fMascaraData(LiberacaoDigitacao.DataFinal) & " Per:" & LiberacaoDigitacao.PeriodoInicial & " a " & LiberacaoDigitacao.PeriodoFinal & " Tp.Mov:" & cboTipoMovimento.Text)
        AtualTabe
        Call GravaAuditoria(1, Me.name, 10, "Para: Dt I:" & fMascaraData(LiberacaoDigitacao.DataInicial) & " - " & fMascaraData(LiberacaoDigitacao.DataFinal) & " Per:" & LiberacaoDigitacao.PeriodoInicial & " a " & LiberacaoDigitacao.PeriodoFinal & " Tp.Mov:" & cboTipoMovimento.Text)
        If Not LiberacaoDigitacao.Alterar(g_empresa, Val(cboTipoMovimento.Text)) Then
            MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
        End If
    End If
    If LiberacaoDigitacao.LocalizarCodigo(g_empresa, Val(cboTipoMovimento.Text)) Then
        AtualTela
        
        If ConfiguracaoDiversa.LocalizarCodigo(1, "PETROMOVELAUTO AUTORIZA NFCE") Then
            
            If ConfiguracaoDiversa.Verdadeiro Then
                g_cfg_periodo_i = LiberacaoDigitacao.PeriodoInicial
                g_cfg_periodo_f = LiberacaoDigitacao.PeriodoFinal
                g_cfg_data_i = LiberacaoDigitacao.DataInicial
                g_cfg_data_f = LiberacaoDigitacao.DataFinal
                
                AtualizaVariaveisGlobaisPetromovelAuto
            End If
        End If
    Else
        MsgBox "Registro inexistente!", vbInformation, "Erro de Verificação!"
    End If
    cmd_sair.SetFocus
    Exit Sub
FileError:
    MsgBox Error
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data_i.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f.Text) < CDate(msk_data_i.Text) Then
        MsgBox "Data final deve ser maior ou igual a " & CDate(msk_data_i) & ".", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf cbo_periodo_i.ListIndex = -1 Then
        MsgBox "Selecione o período inicial.", vbInformation, "Atenção!"
        cbo_periodo_i.SetFocus
    ElseIf cbo_periodo_f.ListIndex = -1 Then
        MsgBox "Selecione o período final.", vbInformation, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cbo_periodo_f < cbo_periodo_i Then
        MsgBox "Periodo final deve ser maior.", vbInformation, "Atenção!"
        cbo_periodo_f.SetFocus
    Else
        ValidaCampos = True
    End If
End Function


Private Sub cmd_sair_Click()
    Finaliza
    Unload Me
End Sub
Private Sub Form_Activate()
    Dim xDados As String

    xDados = ReadINI("CUPOM FISCAL", "Tipo de Venda", gArquivoIni)
    If xDados = "CONVENIENCIA" Then
        cboTipoMovimento.ListIndex = 2
    ElseIf xDados = "AUTOMACAO" Or xDados = "CUPOM FISCAL" Or xDados = "CUPOM FISCAL/CONVENIENCIA" Then
        cboTipoMovimento.ListIndex = 1
    End If

    If LiberacaoDigitacao.LocalizarCodigo(g_empresa, Val(cboTipoMovimento.Text)) Then
        AtualTela
    Else
        CriaNovoRegistro
    End If
    msk_data_i.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_sair_Click
    ElseIf KeyCode = vbKeyF11 Then
        KeyCode = 0
        cmd_ok_Click
    ElseIf KeyCode = vbKeyF12 Then
        KeyCode = 0
        cmd_cancelar_Click
    End If
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    
    PreencheCboTipoMovimento
    PreencheCboPeriodo
    Call GravaAuditoria(1, Me.name, 1, "")
End Sub
Private Sub PreencheCboPeriodo()
    Dim i As Integer
    cbo_periodo_i.Clear
    cbo_periodo_f.Clear
    If Configuracao.LocalizarCodigo(g_empresa) Then
        For i = 1 To Configuracao.QuantidadePeriodos
            cbo_periodo_i.AddItem i
            cbo_periodo_f.AddItem i
            cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = i
            cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = i
        Next
    End If
End Sub
Private Sub PreencheCboTipoMovimento()
    cboTipoMovimento.Clear
    cboTipoMovimento.AddItem "1 Escritório"
    cboTipoMovimento.ItemData(cboTipoMovimento.NewIndex) = 1
    cboTipoMovimento.AddItem "2 Pista"
    cboTipoMovimento.ItemData(cboTipoMovimento.NewIndex) = 2
    cboTipoMovimento.AddItem "3 Conveniência"
    cboTipoMovimento.ItemData(cboTipoMovimento.NewIndex) = 3
End Sub
Private Sub msk_data_f_GotFocus()
    msk_data_f.SelStart = 0
    msk_data_f.SelLength = 2
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_i.SetFocus
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
Private Sub msk_data_i_LostFocus()
    If IsDate(msk_data_i.Text) Then
        msk_data_f.Text = msk_data_i.Text
    End If
End Sub
