VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_identificacao 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Identificação de Usuário"
   ClientHeight    =   2415
   ClientLeft      =   2895
   ClientTop       =   3285
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
   Icon            =   "SENHA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2415
   ScaleWidth      =   6375
   Begin VB.TextBox txt_senha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   900
      Width           =   1215
   End
   Begin VB.ComboBox cbo_usuario 
      Height          =   300
      ItemData        =   "SENHA.frx":0442
      Left            =   2220
      List            =   "SENHA.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   4035
   End
   Begin VB.TextBox txt_codigo_usuario 
      Height          =   285
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin Threed.SSCommand cmd_cancela 
      Cancel          =   -1  'True
      Height          =   825
      Left            =   4500
      TabIndex        =   8
      Top             =   1500
      Width           =   960
      _Version        =   65536
      _ExtentX        =   1693
      _ExtentY        =   1455
      _StockProps     =   78
      Caption         =   "Cancela"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      RoundedCorners  =   0   'False
      Picture         =   "SENHA.frx":0446
   End
   Begin Threed.SSCommand cmd_ok 
      Height          =   825
      Left            =   1080
      TabIndex        =   7
      Tag             =   "b"
      Top             =   1500
      Width           =   960
      _Version        =   65536
      _ExtentX        =   1693
      _ExtentY        =   1455
      _StockProps     =   78
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      RoundedCorners  =   0   'False
      Picture         =   "SENHA.frx":0760
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4680
      Picture         =   "SENHA.frx":0A7A
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lbl_nivel_acesso 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "&Senha:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   900
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Nível de Acesso:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
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
'   Bio-System Informática.
'
'------------------------------------------------------------------
Option Explicit

Dim limite_tentativas As Integer
Const ASC_a = 97, ASC_z = 122
Const ASC_0 = 48, ASC_9 = 57
Const ASC_Espaço = 32




Private Sub Finaliza()
    Dim flag As Integer
    If tbl_usuario!codigo = 99 Then
        flag = 1
    End If
    tbl_usuario.Close
    Unload Me

End Sub

Private Sub MostraNivelAcesso()
    If tbl_usuario!tipo_acesso = 1 Then
        lbl_nivel_acesso.Caption = "Desenvolvimento"
    ElseIf tbl_usuario!tipo_acesso = 2 Then
        lbl_nivel_acesso.Caption = "Diretoria      "
    ElseIf tbl_usuario!tipo_acesso = 3 Then
        lbl_nivel_acesso.Caption = "Gerência       "
    ElseIf tbl_usuario!tipo_acesso = 4 Then
        lbl_nivel_acesso.Caption = "Operação       "
    ElseIf tbl_usuario!tipo_acesso = 5 Then
        lbl_nivel_acesso.Caption = "Digitação      "
    End If

End Sub

Private Sub cbo_usuario_Click()
    If cbo_usuario.ListIndex <> -1 Then
        txt_codigo_usuario = Format(cbo_usuario.ItemData(cbo_usuario.ListIndex), "##")
    Else
        txt_codigo_usuario = ""
    End If

End Sub


Private Sub cbo_usuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_senha.SetFocus
    End If

End Sub



Private Sub cbo_usuario_LostFocus()
    If cbo_usuario.ListIndex <> -1 Then
        tbl_usuario.Seek "=", cbo_usuario.ItemData(cbo_usuario.ListIndex)
        If Not tbl_usuario.NoMatch Then
            MostraNivelAcesso
        Else
            MsgBox "Usuário não cadastrado.", 48, "Atenção!"
            cbo_usuario.ListIndex = -1
            lbl_nivel_acesso = ""
            cbo_usuario.SetFocus
        End If
    End If

End Sub

Private Sub cmd_cancela_Click()
    txt_senha = ""
    txt_senha.SetFocus
    
End Sub

Private Sub cmd_ok_Click()
    If cbo_usuario.ListIndex <> -1 Then
        limite_tentativas = limite_tentativas + 1
        txt_senha = Kriptografa(txt_senha)
        If txt_senha = tbl_usuario!senha Then
            g_usuario = tbl_usuario!codigo
            g_nivel_acesso = tbl_usuario!tipo_acesso
            If g_usuario = 3 Then
                BaixaVencimentos Date
                ScpVbCobol
            End If
            Finaliza
        Else
            If limite_tentativas = 3 Then
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


Private Sub Form_Activate()
    PreencheCboUsuarios
    'If tbl_usuario.RecordCount = 0 Then
    '    DesativaBotoes
    '    cmd_novo.Enabled = True
    '    cmd_sair.Enabled = True
    'Else
    '    AtivaBotoes
    '    BuscaDados
    'End If
    'cmd_novo.SetFocus

End Sub

Private Sub PreencheCboUsuarios()
    cbo_usuario.Clear
    tbl_usuario.Index = "id_nome"
    If tbl_usuario.RecordCount > 0 Then
        tbl_usuario.MoveFirst
        Do Until tbl_usuario.EOF
            cbo_usuario.AddItem tbl_usuario!Nome
            cbo_usuario.ItemData(cbo_usuario.NewIndex) = tbl_usuario!codigo
            tbl_usuario.MoveNext
        Loop
    End If
    tbl_usuario.Index = "id_codigo"
    
End Sub

 Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    Set tbl_usuario = bd_sgp.OpenTable("usuarios")
    tbl_usuario.Index = "id_codigo"

End Sub



Private Sub txt_codigo_usuario_GotFocus()
    txt_codigo_usuario.SelStart = 0
    txt_codigo_usuario.SelLength = Len(txt_codigo_usuario)
    
End Sub

Private Sub txt_codigo_usuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_senha.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub

Private Sub txt_codigo_usuario_LostFocus()
    Dim i As Integer
    If Val(txt_codigo_usuario) > 0 Then
        tbl_usuario.Seek "=", Val(txt_codigo_usuario)
        If Not tbl_usuario.NoMatch Then
            For i = 0 To cbo_usuario.ListCount - 1
                cbo_usuario.ListIndex = i
                If cbo_usuario.ItemData(i) = tbl_usuario!codigo Then
                    MostraNivelAcesso
                    Exit For
                End If
            Next
        Else
            MsgBox "Usuário não cadastrado.", 48, "Atenção!"
            cbo_usuario.ListIndex = -1
            lbl_nivel_acesso = ""
            txt_codigo_usuario.SetFocus
        End If
    Else
        txt_codigo_usuario = ""
        cbo_usuario.SetFocus
    End If

End Sub

Private Sub txt_senha_GotFocus()
    cmd_ok.Default = True

End Sub

Private Sub txt_senha_LostFocus()
    cmd_ok.Default = False

End Sub


