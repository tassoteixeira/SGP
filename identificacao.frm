VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frm_identificacao 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Identificação de Usuário"
   ClientHeight    =   2445
   ClientLeft      =   2475
   ClientTop       =   2250
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "identificacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2445
   ScaleWidth      =   6375
   Begin VB.Timer TimerIdentFid 
      Enabled         =   0   'False
      Left            =   3420
      Top             =   1680
   End
   Begin MSCommLib.MSComm MSCommIdentFid 
      Left            =   2640
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmd_cancelar 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   4620
      Picture         =   "identificacao.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cancela as informações informadas."
      Top             =   1500
      Width           =   855
   End
   Begin VB.CommandButton cmd_ok 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Ok"
      Height          =   855
      Left            =   1080
      Picture         =   "identificacao.frx":193C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Confirma as informações informadas."
      Top             =   1500
      Width           =   855
   End
   Begin VB.TextBox txt_senha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   900
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adodc_usuario 
      Height          =   330
      Left            =   3240
      Top             =   60
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
      Caption         =   "adodc_usuario"
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
   Begin MSDataListLib.DataCombo dtcbo_usuario 
      Bindings        =   "identificacao.frx":2F46
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   60
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "Nome"
      BoundColumn     =   "Codigo"
      Text            =   "dtcbo_usuario"
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4680
      Picture         =   "identificacao.frx":2F62
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lbl_nivel_acesso 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Senha:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   900
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nível de Acesso:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Usuário.:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   6420
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6420
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "frm_identificacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' SENHA.FRM -------------------------------------------------------
'
'   SGP® - Formulário de Entrada de Senha
'
'   © 1996 by Tasso Teixeira
'   Cerrado Informática.
'
'------------------------------------------------------------------
Option Explicit
Dim limite_tentativas As Integer
Const ASC_a As Integer = 97
Const ASC_z As Integer = 122
Const ASC_0 As Integer = 48
Const ASC_9 As Integer = 57
Const ASC_Espaço As Integer = 32

Private CartaoAbastecimento As New cCartaoAbastecimento
Private Empresa As New cEmpresa
Private Funcionario As New cFuncionario
Private Usuario As New cUsuario
Dim lPortaRfid As Integer

Private Sub Finaliza()
    Dim Flag As Integer
    If Usuario.Codigo = 99 Then
        Flag = 1
    End If
    Set CartaoAbastecimento = Nothing
    Set Empresa = Nothing
    Set Funcionario = Nothing
    Set Usuario = Nothing
    Unload Me
End Sub
Private Sub MostraNivelAcesso()
    If Usuario.TipoAcesso = 1 Then
        lbl_nivel_acesso.Caption = "Desenvolvimento"
    ElseIf Usuario.TipoAcesso = 2 Then
        lbl_nivel_acesso.Caption = "Diretoria      "
    ElseIf Usuario.TipoAcesso = 3 Then
        lbl_nivel_acesso.Caption = "Gerência       "
    ElseIf Usuario.TipoAcesso = 4 Then
        lbl_nivel_acesso.Caption = "Operação       "
    ElseIf Usuario.TipoAcesso = 5 Then
        lbl_nivel_acesso.Caption = "Digitação      "
    End If
End Sub
Private Sub cmd_cancelar_Click()
    Dim xNomeArquivo As String
    If (MsgBox("Deseja realmente sair do sistema?", vbYesNo + vbQuestion + vbDefaultButton2, "Sair do Sistema!")) = vbYes Then
        Call GravaAuditoria(1, Me.name, 11, "S.G.P.")
        xNomeArquivo = "C:" & gDiretorioAplicativo & "sgp_cadastro.ini"
        If ReadINI("SGP_CADASTRO", "Sgp_cadastro compilado no VB.NET", gArquivoIni) = "SIM" Then
            Call WriteINI("TIPO DE OPERACAO", "Tipo de Operacao", "Finaliza SGP", xNomeArquivo)
        End If
        If gArqTxt.FolderExists("C:\Cerrado.Net\SgpNet") Then
            Call WriteINI("TIPO DE OPERACAO", "Tipo de Operacao", "Finaliza SGP", "C:\Cerrado.Net\SgpNet\SgpNetTemporario.ini")
        End If
        End
    Else
        txt_senha.Text = ""
        txt_senha.SetFocus
    End If
End Sub
Private Sub cmd_ok_Click()
    Dim x_resultado As Integer
    
    If dtcbo_usuario.BoundText <> "" Then
        limite_tentativas = limite_tentativas + 1
        'MsgBox DesKriptografa(Usuario.Senha)
        If Kriptografa(txt_senha.Text) = Usuario.Senha Then
            limite_tentativas = 0
            g_usuario = Usuario.Codigo
            g_nome_usuario = Usuario.Nome
            g_nivel_acesso = Usuario.TipoAcesso
'            g_situacao = Usuario.Situacao
            Call GravaAuditoria(1, Me.name, 16, "")
            If Empresa.LocalizarPrimeiro Then
                'If UCase(Empresa.Nome) = "MARQUES DE CASTRO & GABRIEL LTDA" Then
                '    If (MsgBox("Deseja fazer o sincronismo com a Automação?", vbDefaultButton2 + vbQuestion + vbYesNo, "Sincronismo com Automação!")) = 6 Then
                '        x_resultado = SincronismoAutomacao
                '        If x_resultado = 0 Then
                '            MsgBox "Sincronismo feito com sucesso!"
                '        Else
                '            MsgBox "Erro de Sincronismo: " & x_resultado
                '        End If
                '    End If
                'End If
            End If
'            If Mid(Usuario.Nome, 1, 9) = "Sonia" Then
'                If Day(g_data_def) >= 16 And Day(g_data_def) <= 21 Then
'                    Do Until (MsgBox("Não esqueça do fechamento da AFOJAC." & Chr(10) & Chr(10) & "Escolha OK para continuar.", 1 + 256, "Mensagem para o Operador de Sistema")) = 1
'                    Loop
'                End If
'                If (Day(g_data_def) >= 7 And Day(g_data_def) <= 9) Or (Day(g_data_def) >= 17 And Day(g_data_def) <= 19) Or (Day(g_data_def) >= 27 And Day(g_data_def) <= 29) Then
'                    Do Until (MsgBox("Não esqueça do fechamento da Goiânia Transportes." & Chr(10) & Chr(10) & "Escolha OK para continuar.", 1 + 256, "Mensagem para o Operador de Sistema")) = 1
'                    Loop
'                End If
'                'If Day(g_data_def) >= 20 And Day(g_data_def) <= 26 Then
'                '    Do Until (MsgBox("Não esqueça do fechamento da ASSEGO." & Chr(10) & Chr(10) & "Escolha OK para continuar.", 1 + 256, "Mensagem para o Operador de Sistema")) = 1
'                '    Loop
'                'End If
'            End If
            Finaliza
        Else
            Call GravaAuditoria(1, Me.name, 17, Usuario.Nome)
            If limite_tentativas = 3 Then
                limite_tentativas = 0
                MsgBox "Acesso não autorizado." & vbCrLf & "Este aplicativo será fechado.", vbInformation, "Senha Inválida!"
                Finaliza
            Else
                MsgBox "Senha informada não confere." & vbCrLf & "Informe pela " & limite_tentativas + 1 & "a vez.", vbInformation, "Senha Inválida!"
                txt_senha.Text = ""
                txt_senha.SetFocus
            End If
        End If
    End If
End Sub
Private Sub dtcbo_usuario_GotFocus()
    'MsgBox "lPortaRfid=" & lPortaRfid
    If lPortaRfid > 0 Then
        If MSCommIdentFid.PortOpen = False Then
            'MsgBox "Configura e abre porta"
            MSCommIdentFid.CommPort = lPortaRfid
            MSCommIdentFid.Settings = "9600,n,8,1"
            MSCommIdentFid.PortOpen = True
            MSCommIdentFid.InBufferCount = 0
            TimerIdentFid.Enabled = True
            TimerIdentFid.Interval = 500
            'MsgBox "Porta MSCommIdentFid.PortOpen=" & MSCommIdentFid.PortOpen
        End If
    End If
End Sub
Private Sub dtcbo_usuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_senha.SetFocus
    End If
End Sub
Private Sub dtcbo_usuario_LostFocus()
    If MSCommIdentFid.PortOpen = True Then
        MSCommIdentFid.PortOpen = False
        TimerIdentFid.Enabled = False
    End If
    If dtcbo_usuario.BoundText <> "" Then
        If Usuario.LocalizarCodigo(Val(dtcbo_usuario.BoundText)) Then
            MostraNivelAcesso
            If Usuario.Situacao = "I" Then
                If (MsgBox("Atenção!" & vbCrLf & "Este usuário está inativo" & vbCrLf & "Deseja continuar?", vbYesNo + vbQuestion + vbDefaultButton2, "Usuário Inativo!")) = vbNo Then
                    dtcbo_usuario.SetFocus
                    Exit Sub
                End If
                cmd_ok.Default = True
            End If
        Else
            MsgBox "Usuário não cadastrado.", vbInformation, "Atenção!"
            dtcbo_usuario.BoundText = ""
            lbl_nivel_acesso.Caption = ""
            dtcbo_usuario.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF11 Then
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
    
'    adodc_usuario.ConnectionString = gConnectionString
'    adodc_usuario.RecordSource = "SELECT Codigo, Nome FROM Usuario WHERE Situacao = " & Chr(39) & "A" & Chr(39) & " ORDER BY Nome"
'    adodc_usuario.Refresh
    Set adodc_usuario.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Usuario WHERE Situacao = " & preparaTexto("A") & " ORDER BY Nome")
    lPortaRfid = 0
    lPortaRfid = Val(ReadINI("CUPOM FISCAL", "Porta Leitor RFID", gArquivoIni))
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then
        If (MsgBox("Deseja realmente sair do sistema?", vbYesNo + vbQuestion + vbDefaultButton2, "Sair do Sistema!")) = vbNo Then
            Cancel = True
        End If
    End If
End Sub
Private Sub txt_senha_GotFocus()
    txt_senha.SelStart = 0
    txt_senha.SelLength = Len(txt_senha.Text)
    cmd_ok.Default = True
End Sub
Private Sub txt_senha_LostFocus()
    cmd_ok.Default = False
End Sub
Private Sub TimerIdentFid_Timer()
    Dim i As Integer
    Dim xString As String
    
    xString = MSCommIdentFid.Input
    If Len(xString) = 23 Then
        MsgBox "xstring=" & xString & vbCrLf & "Cartao=" & Mid(xString, 4, 16) & vbCrLf & "Tamanho=" & Len(xString)
        TimerIdentFid.Enabled = False
        If MSCommIdentFid.PortOpen = True Then
            MSCommIdentFid.PortOpen = False
        End If
        If CartaoAbastecimento.LocalizarNumeroCartao(g_empresa, Mid(xString, 4, 16)) Then
            If Funcionario.LocalizarCodigo(CartaoAbastecimento.Empresa, CartaoAbastecimento.CodigoFuncionario) Then
                dtcbo_usuario.BoundText = Funcionario.CodigoUsuario
                If Usuario.LocalizarCodigo(Funcionario.CodigoUsuario) Then
                    MostraNivelAcesso
                    txt_senha.Text = DesKriptografa(Usuario.Senha)
                    cmd_ok_Click
                Else
                    txt_senha.SetFocus
                End If
            Else
                txt_senha.SetFocus
            End If
        End If
    ElseIf Len(xString) > 0 Then
        MsgBox "xstring=" & xString & vbCrLf & "Cartao=" & Mid(xString, 4, 16) & vbCrLf & "Tamanho=" & Len(xString)
    End If
End Sub

