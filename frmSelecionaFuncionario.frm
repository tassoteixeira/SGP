VERSION 5.00
Begin VB.Form frmSelecionaFuncionario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecione o funcionário"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   6000
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtCodigoFuncionario 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cboFuncionario 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmSelecionaFuncionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lFuncionario As New CadastroDLL.cFuncionario
Private rstFuncionario As adodb.Recordset

Public gCodigoFuncionarioSelecionado As Integer
Public gNomeFuncionarioSelecionado As String


Private Function ValidaCampos() As Boolean

ValidaCampos = False

If Len(Trim(txtCodigoFuncionario.Text)) = 0 Then
    MsgBox "O Codigo do Funcionário deve ser informado", vbInformation, "Campo obrigatório"
    txtCodigoFuncionario.SetFocus
ElseIf cboFuncionario.ListIndex <= 0 Then
    MsgBox "O Funcionário deve ser selecionado", vbInformation, "Campo obrigatório"
    cboFuncionario.SetFocus
Else
    ValidaCampos = True
End If

End Function

Private Sub btnCancelar_Click()
    gCodigoFuncionarioSelecionado = 0
    gNomeFuncionarioSelecionado = ""
    Unload Me
End Sub

Private Sub btnOK_Click()
    
    If ValidaCampos Then
        gCodigoFuncionarioSelecionado = Val(txtCodigoFuncionario.Text)
        gNomeFuncionarioSelecionado = cboFuncionario.Text
                
        Unload Me
    End If
End Sub

Private Sub cboFuncionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        btnOK.SetFocus
    End If
End Sub

Private Sub cboFuncionario_LostFocus()
    txtCodigoFuncionario.Text = IIf(cboFuncionario.ItemData(cboFuncionario.ListIndex) = 0, "", CStr(cboFuncionario.ItemData(cboFuncionario.ListIndex)))
End Sub

Private Sub Form_Activate()
    txtCodigoFuncionario.SetFocus
End Sub

Private Sub txtCodigoFuncionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        btnOK.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub

Private Sub txtCodigoFuncionario_LostFocus()
    Dim i As Integer
    cboFuncionario.ListIndex = 0
    
    For i = 0 To cboFuncionario.ListCount - 1
        If cboFuncionario.ItemData(i) = Val(txtCodigoFuncionario.Text) Then
            cboFuncionario.ListIndex = i
            Exit Sub
        End If
    Next
    
    MsgBox "Funcionário não encontrado", vbExclamation, "Registro não encontrado"
    
End Sub


Private Sub Form_Load()
    PreencheCboFuncionario
   
    gCodigoFuncionarioSelecionado = 0
    gNomeFuncionarioSelecionado = ""

End Sub

Private Sub PreencheCboFuncionario()

    Dim xSQL As String
    
    cboFuncionario.Clear
   
    cboFuncionario.AddItem ""
    cboFuncionario.ItemData(cboFuncionario.NewIndex) = 0
   
    xSQL = "SELECT Codigo, Nome FROM Funcionario ORDER BY Nome"
    Set rstFuncionario = Conectar.RsConexao(xSQL)
    With rstFuncionario
        If .RecordCount > 0 Then
            Do Until .EOF
                cboFuncionario.AddItem !Nome
                cboFuncionario.ItemData(cboFuncionario.NewIndex) = !Codigo
                .MoveNext
            Loop
            rstFuncionario.Close
        End If
    End With
    
    cboFuncionario.ListIndex = 0
    Set rstFuncionario = Nothing

End Sub

