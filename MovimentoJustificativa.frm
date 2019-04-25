VERSION 5.00
Begin VB.Form MovimentoJustificativa 
   Caption         =   "Justificativa"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8580
   Icon            =   "MovimentoJustificativa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmDados 
      Height          =   1155
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8355
      Begin VB.TextBox txtJustificativa 
         Height          =   315
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   4
         Top             =   660
         Width           =   6675
      End
      Begin VB.Label lblOperacao 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label3 
         Caption         =   "Operação"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Justificativa"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   2400
      Picture         =   "MovimentoJustificativa.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Confirma o registro atual."
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5400
      Picture         =   "MovimentoJustificativa.frx":1914
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1320
      Width           =   795
   End
End
Attribute VB_Name = "MovimentoJustificativa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lString As String
Dim lCodigoFuncionario As Integer

Private MovJustificativa As New cMovimentoJustificativa
Private Sub Form_Activate()
    Screen.MousePointer = 1
    lString = g_string
    g_string = ""
    lblOperacao.Caption = RetiraString(1, lString)
    lCodigoFuncionario = Val(RetiraString(3, lString))
    txtJustificativa.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF11 Then
        KeyCode = 0
        cmd_ok_Click
    End If
End Sub
Private Sub Form_Load()
    Call GravaAuditoria(1, Me.name, 1, "")
    CentraForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    
    If ValidaCampos Then
        AtualizaTabela
        Call GravaAuditoria(1, Me.name, 10, "Operacao:" & lblOperacao.Caption)
        Call GravaAuditoria(2, Me.name, 10, " Justif:" & txtJustificativa.Text)
        If MovJustificativa.Incluir Then
            g_string = "OK|@|" & MovJustificativa.numero & "|@|"
            cmd_sair_Click
        Else
            MsgBox "Não foi possível incluir o registro de justificativa!", vbCritical, "Erro de Integridade!"
            cmd_sair_Click
        End If
    End If
    Exit Sub
    
FileError:
    MsgBox "Erro desconhecido na atualização de registro!", vbInformation, "Erro de Integridade"
    Exit Sub
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub txtJustificativa_GotFocus()
    txtJustificativa.SelStart = 0
    txtJustificativa.SelLength = Len(txtJustificativa.Text)
End Sub
Private Sub txtJustificativa_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub AtualizaTabela()
    MovJustificativa.numero = 1
    MovJustificativa.Data = Format(Now, "dd/mm/yyyy")
    MovJustificativa.Hora = Format(Now, "HH:mm:ss")
    MovJustificativa.Operacao = lblOperacao.Caption
    MovJustificativa.CodigoUsuario = g_usuario
    MovJustificativa.NomeInternoPrograma = RetiraString(2, lString)
    MovJustificativa.Computador = GetIPHostName() & " - " & GetIPAddress()
    MovJustificativa.Justificativa = txtJustificativa.Text
    MovJustificativa.CodigoFuncionario = lCodigoFuncionario
    MovJustificativa.DadosInterno = ""
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set MovJustificativa = Nothing
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Trim(txtJustificativa.Text) = "" Then
        MsgBox "Informe a justificativa desta operação.", vbInformation, "Atenção!"
        txtJustificativa.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
