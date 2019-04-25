VERSION 5.00
Begin VB.Form frmMensagemAutomatica 
   Caption         =   "Título da Mensagem"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   3780
      Top             =   2400
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK (01)"
      Height          =   435
      Left            =   1680
      TabIndex        =   0
      Top             =   2460
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   60
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblMensagem 
      Caption         =   "Corpo da mensagem"
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   4395
   End
End
Attribute VB_Name = "frmMensagemAutomatica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lQtdSegundos As Integer

Private Sub cmdOk_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    CentraForm Me
    Me.Caption = RetiraString(1, gString)
    lblMensagem.Caption = RetiraString(2, gString)
    lQtdSegundos = Val(RetiraString(3, gString))
    cmdOk.Caption = "&Ok (" & Format(lQtdSegundos) & ")"
    Timer1.Interval = 1000
    Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
    lQtdSegundos = lQtdSegundos - 1
    cmdOk.Caption = "&Ok (" & Format(lQtdSegundos) & ")"
    DoEvents
    'MsgBox "TEF Confirmado com sucesso!", vbInformation, "TEF Concluído"
    If lQtdSegundos = 0 Then
        cmdOk_Click
    End If
End Sub
