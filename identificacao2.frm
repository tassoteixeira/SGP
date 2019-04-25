VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frm_identificacao2 
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
   Icon            =   "identificacao2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2445
   ScaleWidth      =   6375
   Begin VB.CommandButton cmd_cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   4620
      Picture         =   "identificacao2.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cancela as informações informadas."
      Top             =   1500
      Width           =   855
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   1080
      Picture         =   "identificacao2.frx":171C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Confirma as informações informadas."
      Top             =   1500
      Width           =   855
   End
   Begin VB.Data dta_usuario 
      Caption         =   "dta_usuario"
      Connect         =   "Access"
      DatabaseName    =   "\VB5\SGP\DATA\Sgp_data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3540
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Usuario"
      Top             =   60
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txt_senha 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   900
      Width           =   1215
   End
   Begin MSDBCtls.DBCombo dbcbo_usuario 
      Bindings        =   "identificacao2.frx":29F6
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   60
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      _Version        =   327680
      MatchEntry      =   -1  'True
      Style           =   2
      BackColor       =   16777215
      ForeColor       =   0
      ListField       =   "Nome"
      BoundColumn     =   "Codigo"
      Text            =   ""
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4680
      Picture         =   "identificacao2.frx":2A13
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
      Caption         =   "&Senha"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   900
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Nível de Acesso"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "&Usuário"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   60
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
Attribute VB_Name = "frm_identificacao2"
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
Dim tbl_usuario As Table
Dim limite_tentativas As Integer
Const ASC_a = 97, ASC_z = 122
Const ASC_0 = 48, ASC_9 = 57
Const ASC_Espaço = 32
Private Sub Finaliza()
    Dim flag As Integer
    If tbl_usuario!Codigo = 99 Then
        flag = 1
    End If
    tbl_usuario.Close
    Unload Me
End Sub
Private Sub MostraNivelAcesso()
    If tbl_usuario![Tipo de Acesso] = 1 Then
        lbl_nivel_acesso.Caption = "Desenvolvimento"
    ElseIf tbl_usuario![Tipo de Acesso] = 2 Then
        lbl_nivel_acesso.Caption = "Diretoria      "
    ElseIf tbl_usuario![Tipo de Acesso] = 3 Then
        lbl_nivel_acesso.Caption = "Gerência       "
    ElseIf tbl_usuario![Tipo de Acesso] = 4 Then
        lbl_nivel_acesso.Caption = "Operação       "
    ElseIf tbl_usuario![Tipo de Acesso] = 5 Then
        lbl_nivel_acesso.Caption = "Digitação      "
    End If
End Sub
Private Sub cmd_cancelar_Click()
    If (MsgBox("Deseja realmente sair do sistema?", 4 + 32 + 256, "Sair do Sistema!")) = 6 Then
        End
    Else
        txt_senha = ""
        txt_senha.SetFocus
    End If
End Sub
Private Sub cmd_ok_Click()
    If Val(dbcbo_usuario.BoundText) > 0 Then
        limite_tentativas = limite_tentativas + 1
        txt_senha = Kriptografa(txt_senha)
        'txt_senha = DesKriptografa(tbl_usuario!Senha)
        If txt_senha = tbl_usuario!Senha Then
            limite_tentativas = 0
            g_usuario = tbl_usuario!Codigo
            g_nome_usuario = tbl_usuario!Nome
            g_nivel_acesso = tbl_usuario![Tipo de Acesso]
'            g_situacao = tbl_usuario!Situacao
            If g_usuario = 3 And Mid(tbl_usuario!Nome, 1, 7) = "Roberto" Then
                BaixaVencimentos Date
'                ScpVbCobol
'                DuplicataReceberVbCobol
            ElseIf g_usuario = 5 Then
                BaixaChequePreDatado Date
            End If
            If Mid(tbl_usuario!Nome, 1, 9) = "Sonia" Then
                If Day(g_data_def) >= 16 And Day(g_data_def) <= 21 Then
                    Do Until (MsgBox("Não esqueça do fechamento da AFOJAC." & Chr(10) & Chr(10) & "Escolha OK para continuar.", 1 + 256, "Mensagem para o Operador de Sistema")) = 1
                    Loop
                End If
                If (Day(g_data_def) >= 7 And Day(g_data_def) <= 9) Or (Day(g_data_def) >= 17 And Day(g_data_def) <= 19) Or (Day(g_data_def) >= 27 And Day(g_data_def) <= 29) Then
                    Do Until (MsgBox("Não esqueça do fechamento da Goiânia Transportes." & Chr(10) & Chr(10) & "Escolha OK para continuar.", 1 + 256, "Mensagem para o Operador de Sistema")) = 1
                    Loop
                End If
                'If Day(g_data_def) >= 20 And Day(g_data_def) <= 26 Then
                '    Do Until (MsgBox("Não esqueça do fechamento da ASSEGO." & Chr(10) & Chr(10) & "Escolha OK para continuar.", 1 + 256, "Mensagem para o Operador de Sistema")) = 1
                '    Loop
                'End If
            End If
            Finaliza
        Else
            If limite_tentativas = 3 Then
                limite_tentativas = 0
                MsgBox "Acesso não autorizado." & Chr(10) & "Este aplicativo será fechado.", 64, "Senha Inválida!"
                Finaliza
            Else
                MsgBox "Senha informada não confere." & Chr(10) & "Informe pela " & limite_tentativas + 1 & "a vez.", 64, "Senha Inválida!"
                txt_senha = ""
                txt_senha.SetFocus
            End If
        End If
    End If
End Sub
Private Sub dbcbo_usuario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_senha.SetFocus
    End If
End Sub
Private Sub dbcbo_usuario_LostFocus()
    If Val(dbcbo_usuario.BoundText) > 0 Then
        tbl_usuario.Seek "=", Val(dbcbo_usuario.BoundText)
        If Not tbl_usuario.NoMatch Then
            MostraNivelAcesso
            If tbl_usuario!Situacao = "I" Then
                If (MsgBox("Atenção!" & Chr(10) & "Este usuário está inativo" & Chr(10) & "Deseja continuar?", 4 + 32 + 256, "Usuário Inativo!")) = 7 Then
                    dbcbo_usuario.SetFocus
                    Exit Sub
                End If
                cmd_ok.Default = True
            End If
        Else
            MsgBox "Usuário não cadastrado.", 48, "Atenção!"
            dbcbo_usuario.BoundText = 0
            lbl_nivel_acesso = ""
            dbcbo_usuario.SetFocus
        End If
    End If
End Sub
Private Sub Form_Activate()
    AtualizaUsuarios
End Sub
Private Sub AtualizaUsuarios()
    dta_usuario.RecordSource = "Select * From Usuario Order By Nome"
    dta_usuario.Refresh
End Sub
 Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    Set tbl_usuario = bd_sgp.OpenTable("Usuario")
    tbl_usuario.Index = "id_codigo"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then
        If (MsgBox("Deseja realmente sair do sistema?", 4 + 32 + 256, "Sair do Sistema!")) = 7 Then
            Cancel = True
        End If
    End If
End Sub
Private Sub txt_senha_GotFocus()
    txt_senha.SelStart = 0
    txt_senha.SelLength = Len(txt_senha)
    cmd_ok.Default = True
End Sub
Private Sub txt_senha_LostFocus()
    cmd_ok.Default = False
End Sub
