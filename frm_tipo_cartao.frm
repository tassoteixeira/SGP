VERSION 5.00
Begin VB.Form frm_tipo_cartao 
   Caption         =   "Tipo de Cartão"
   ClientHeight    =   555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   555
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_ok2 
      Caption         =   "O&K"
      Height          =   375
      Left            =   3420
      TabIndex        =   1
      ToolTipText     =   "Confirma o cartão selecionado."
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox cbo_tipo_cartao 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3195
   End
End
Attribute VB_Name = "frm_tipo_cartao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_tipo_cartao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok2_Click
    End If
End Sub
Private Sub cmd_ok2_Click()
    If cbo_tipo_cartao.ListIndex <> -1 Then
        g_string = cbo_tipo_cartao.Text
        Unload Me
    Else
        MsgBox "Selecione um cartão!", vbInformation, "Atenção!"
        cbo_tipo_cartao.SetFocus
    End If
End Sub
Private Sub Form_Load()
    Call PreencheTipoCartao
End Sub
Private Sub PreencheTipoCartao()
    cbo_tipo_cartao.Clear
    cbo_tipo_cartao.AddItem "AMEX"
    cbo_tipo_cartao.ItemData(cbo_tipo_cartao.NewIndex) = 1
    cbo_tipo_cartao.AddItem "REDECARD"
    cbo_tipo_cartao.ItemData(cbo_tipo_cartao.NewIndex) = 2
    cbo_tipo_cartao.AddItem "TECBAN"
    cbo_tipo_cartao.ItemData(cbo_tipo_cartao.NewIndex) = 3
    cbo_tipo_cartao.AddItem "VISANET"
    cbo_tipo_cartao.ItemData(cbo_tipo_cartao.NewIndex) = 4
    cbo_tipo_cartao.AddItem "HIPERTEF"
    cbo_tipo_cartao.ItemData(cbo_tipo_cartao.NewIndex) = 5
End Sub

