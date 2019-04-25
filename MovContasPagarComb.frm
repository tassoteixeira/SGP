VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form MovContasPagarComb 
   Caption         =   "Form1"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5520
      Picture         =   "MovContasPagarComb.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Confirma o registro atual."
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6420
      Picture         =   "MovContasPagarComb.frx":160A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cancela o registro atual."
      Top             =   1800
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton cmd_data_v 
         Height          =   315
         Left            =   2880
         Picture         =   "MovContasPagarComb.frx":2B04
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_complemento 
         Height          =   285
         Left            =   1740
         MaxLength       =   40
         TabIndex        =   6
         Top             =   960
         Width           =   4935
      End
      Begin VB.TextBox txt_valor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_data_vencimento 
         Height          =   300
         Left            =   1740
         TabIndex        =   2
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   255
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Complemento"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Data do Vencimento"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Valor do Vencimento"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
   End
End
Attribute VB_Name = "MovContasPagarComb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MovContaPagar As New cMovimentoContaPagar
Private Sub AtualTabe()
    lRegistro = Val(txt_registro.Text)
    MovContaPagar.Empresa = g_empresa
    MovContaPagar.Registro = lRegistro
    MovContaPagar.CodigoFornecedor = Val(dtcboFornecedor.BoundText)
    MovContaPagar.NomeFornecedor = dtcboFornecedor
    MovContaPagar.DataEmissao = msk_data_emissao.Text
    MovContaPagar.DataVencimento = msk_data_vencimento.Text
    MovContaPagar.Valor = fValidaValor2(txt_valor.Text)
    MovContaPagar.NumeroDocumento = txt_numero_documento.Text
    MovContaPagar.LocalCobranca = cbo_local.ItemData(cbo_local.ListIndex)
    MovContaPagar.CodigoConta = cbo_conta.ItemData(cbo_conta.ListIndex)
    MovContaPagar.Complemento = txt_complemento.Text
    MovContaPagar.DataDigitacao = CDate(lbl_data_digitacao.Caption)
    MovContaPagar.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
End Sub
Private Sub cmd_data_v_Click()
    g_string = msk_data_emissao
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_emissao = RetiraGString(1)
        msk_data_vencimento = RetiraGString(2)
    Else
        msk_data_vencimento = RetiraGString(1)
    End If
    g_string = " "
    txt_valor.SetFocus
End Sub

Private Sub msk_data_vencimento_GotFocus()
    If Not IsDate(msk_data_vencimento) Then
        msk_data_vencimento = "__/__/" & Format(g_data_def, "yyyy")
    End If
End Sub
Private Sub msk_data_vencimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_vencimento.SetFocus
    End If
End Sub
Private Sub txt_complemento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_valor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_complemento.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_valor_LostFocus()
    txt_valor = Format(txt_valor, "###,##0.00")
End Sub
