VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form MensagemImportante 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensagem Importante!"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdContinuar 
      Caption         =   "&Continuar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9060
      TabIndex        =   3
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CheckBox chkConcordo 
      Caption         =   "Mensagem lida, seguirei conforme a instrução acima."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   6375
   End
   Begin RichTextLib.RichTextBox txtMensagem 
      Height          =   3675
      Left            =   60
      TabIndex        =   1
      Top             =   1200
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6482
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"MensagemImportante.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "Título da mensagem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   10335
   End
End
Attribute VB_Name = "MensagemImportante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Form_Activate()
    Screen.MousePointer = 1
    txtMensagem.SetFocus
End Sub
Private Sub Form_Load()
    CentraForm Me
    Call GravaAuditoria(1, Me.name, 1, "")
    cmdContinuar.Enabled = True
    If Len(g_string) > 0 Then
        cmdContinuar.Enabled = False
        lblTitulo.Caption = RetiraGString(1)
        txtMensagem.Text = RetiraGString(2)
        Call GravaAuditoria(1, Me.name, 22, Mid(txtMensagem.Text, 1, 40))
        g_string = ""
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub cmdContinuar_Click()
    Call GravaAuditoria(1, Me.name, 10, "")
    Unload Me
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
End Sub
Private Sub chkConcordo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If chkConcordo.Value = 1 Then
            cmdContinuar.Enabled = True
            cmdContinuar.SetFocus
        Else
            cmdContinuar.Enabled = False
        End If
    End If
End Sub
Private Sub chkConcordo_LostFocus()
    If chkConcordo.Value = 1 Then
        cmdContinuar.Enabled = True
        cmdContinuar.SetFocus
    Else
        cmdContinuar.Enabled = False
    End If
End Sub
Private Sub chkConcordo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If chkConcordo.Value = 1 Then
        cmdContinuar.Enabled = True
        cmdContinuar.SetFocus
    Else
        cmdContinuar.Enabled = False
    End If
End Sub
