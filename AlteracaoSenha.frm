VERSION 5.00
Begin VB.Form AlteracaoSenha 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alteração de Senha"
   ClientHeight    =   3255
   ClientLeft      =   2475
   ClientTop       =   2250
   ClientWidth     =   7635
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
   Icon            =   "AlteracaoSenha.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3255
   ScaleWidth      =   7635
   Begin VB.TextBox txtSenhaNova2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2700
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtSenhaNova1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2700
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmd_cancelar 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   4800
      Picture         =   "AlteracaoSenha.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cancela as informações informadas."
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmd_ok 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Ok"
      Height          =   855
      Left            =   1980
      Picture         =   "AlteracaoSenha.frx":193C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Confirma as informações informadas."
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtSenhaAtual 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2700
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Confirma Senha Nova"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Senha &Nova"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   2475
   End
   Begin VB.Label lblNomeUsuario 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3360
      TabIndex        =   2
      Top             =   60
      Width           =   4215
   End
   Begin VB.Label lblCodigoUsuario 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2700
      TabIndex        =   1
      Top             =   60
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4800
      Picture         =   "AlteracaoSenha.frx":2F46
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label lblNivelAcesso 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2700
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Senha Atual"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2475
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nível de Acesso"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2475
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Usuário"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2475
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   7620
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   7620
      Y1              =   900
      Y2              =   900
   End
End
Attribute VB_Name = "AlteracaoSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lLimiteTentativas As Integer
Dim lCodigoUsuario As Integer

Const ASC_a = 97, ASC_z = 122
Const ASC_0 = 48, ASC_9 = 57
Const ASC_Espaço = 32

Private Usuario As New cUsuario
Private Sub Finaliza()
    Set Usuario = Nothing
    Unload Me
End Sub
Private Function VerificaSenhaAtual() As Boolean
    VerificaSenhaAtual = False
    
    lLimiteTentativas = lLimiteTentativas + 1
    If Kriptografa(txtSenhaAtual.Text) = Usuario.Senha Then
        lLimiteTentativas = 0
        VerificaSenhaAtual = True
        Call GravaAuditoria(1, Me.name, 20, Usuario.Nome & " - Senha Atual Identificada.")
        Finaliza
    Else
        Call GravaAuditoria(1, Me.name, 17, Usuario.Nome & " - Senha Atual Errada.")
        If lLimiteTentativas = 3 Then
            lLimiteTentativas = 0
            MsgBox "Alteração de senha não autorizada.", vbInformation, "Senha Inválida!"
            Finaliza
        Else
            MsgBox "Senha informada não confere." & Chr(10) & "Informe pela " & lLimiteTentativas + 1 & "a vez.", vbInformation, "Senha Inválida!"
            txtSenhaAtual.Text = ""
        End If
    End If
End Function
Private Function VerificaSenhas() As Boolean
    VerificaSenhas = False
    
    If VerificaSenhaAtual Then
        If txtSenhaNova1.Text = "" Then
            MsgBox "A nova senha não pode ser vazia.", vbInformation, "Dados Inválido!"
            txtSenhaNova1.SetFocus
        ElseIf Len(txtSenhaNova1.Text) < 4 Then
            MsgBox "A nova senha não pode ter menos que 4 caractere.", vbInformation, "Dados Inválido!"
            txtSenhaNova1.SetFocus
        ElseIf txtSenhaNova1.Text = txtSenhaAtual Then
            MsgBox "A nova senha deve ser diferente da senha atual.", vbInformation, "Dados Inválido!"
            txtSenhaNova1.SetFocus
        ElseIf txtSenhaNova1.Text <> txtSenhaNova2.Text Then
            MsgBox "A confirmação da nova senha não confere.", vbInformation, "Dados Inválido!"
            txtSenhaNova2.SetFocus
        Else
            VerificaSenhas = True
        End If
    End If
End Function
Private Sub cmd_cancelar_Click()
    Dim xNomeArquivo As String
    If (MsgBox("Deseja realmente sair do sistema?", 4 + 32 + 256, "Sair do Sistema!")) = 6 Then
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
        txtSenhaAtual.Text = ""
        txtSenhaAtual.SetFocus
    End If
End Sub
Private Sub cmd_ok_Click()
    If VerificaSenhas Then
        Usuario.Senha = Kriptografa(txtSenhaNova1.Text)
        If Usuario.Alterar(lCodigoUsuario) Then
            Call GravaAuditoria(1, Me.name, 10, Usuario.Nome & " - Senha Alterada.")
        Else
            Call GravaAuditoria(1, Me.name, 22, Usuario.Nome & " - Erro ao alterar Senha.")
            MsgBox "Não foi possível alterar a senha.", vbCritical, "Erro de Integridade!"
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
    lLimiteTentativas = 0
    If Len(g_string) > 0 Then
        lCodigoUsuario = RetiraGString(1)
        g_string = ""
        If Usuario.LocalizarCodigo(lCodigoUsuario) Then
            lblCodigoUsuario.Caption = Usuario.Codigo
            lblNomeUsuario.Caption = Usuario.Nome
            lblNivelAcesso.Caption = NivelAcesso(Usuario.TipoAcesso)
        Else
            MsgBox "Usuário inexistente!", vbCritical, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then
        If (MsgBox("Deseja realmente sair do sistema?", 4 + 32 + 256, "Sair do Sistema!")) = 7 Then
            Cancel = True
        End If
    End If
End Sub
Private Sub txtSenhaAtual_GotFocus()
    txtSenhaAtual.SelStart = 0
    txtSenhaAtual.SelLength = Len(txtSenhaAtual.Text)
End Sub
Private Sub txtSenhaAtual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Len(txtSenhaAtual.Text) > 0 Then
            If VerificaSenhaAtual Then
                txtSenhaNova1.SetFocus
            Else
                txtSenhaAtual.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txtSenhaAtual_LostFocus()
    cmd_ok.Default = False
End Sub
Private Sub txtSenhaNova1_GotFocus()
    txtSenhaNova1.SelStart = 0
    txtSenhaNova1.SelLength = Len(txtSenhaNova1.Text)
End Sub
Private Sub txtSenhaNova1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtSenhaNova2.SetFocus
    End If
End Sub
Private Sub txtSenhaNova2_GotFocus()
    txtSenhaNova2.SelStart = 0
    txtSenhaNova2.SelLength = Len(txtSenhaNova2.Text)
End Sub
Private Sub txtSenhaNova2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
