VERSION 5.00
Begin VB.Form frm_grupo 
   Caption         =   "Cadastro de Grupo"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   6120
      TabIndex        =   13
      Top             =   2700
      Width           =   855
   End
   Begin VB.CommandButton cmd_primeiro 
      Caption         =   "&Primeiro"
      Height          =   375
      Left            =   1380
      TabIndex        =   12
      Top             =   2340
      Width           =   1035
   End
   Begin VB.CommandButton cmd_anterior 
      Caption         =   "&Anterior"
      Height          =   375
      Left            =   2580
      TabIndex        =   11
      Top             =   2340
      Width           =   1035
   End
   Begin VB.CommandButton cmd_proximo 
      Caption         =   "&Próximo"
      Height          =   375
      Left            =   3780
      TabIndex        =   10
      Top             =   2340
      Width           =   1035
   End
   Begin VB.CommandButton cmd_ultimo 
      Caption         =   "&Último"
      Height          =   375
      Left            =   4980
      TabIndex        =   9
      Top             =   2340
      Width           =   1035
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5580
      TabIndex        =   8
      Top             =   1740
      Width           =   1035
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   1740
      Width           =   1035
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   1740
      Width           =   1035
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1740
      Width           =   1035
   End
   Begin VB.CommandButton cmd_incluir 
      Caption         =   "&Incluir"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1740
      Width           =   1035
   End
   Begin VB.TextBox txt_nome 
      Height          =   315
      Left            =   1740
      MaxLength       =   30
      TabIndex        =   3
      Top             =   780
      Width           =   5235
   End
   Begin VB.TextBox txt_codigo 
      Height          =   315
      Left            =   1740
      MaxLength       =   3
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   300
      Picture         =   "frm_grupo.frx":0000
      Top             =   2280
      Width           =   525
   End
   Begin VB.Label Label2 
      Caption         =   "Nome do Grupo"
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   780
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Código do Grupo"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   1455
   End
End
Attribute VB_Name = "frm_grupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Grupo As cGrupo

Private Sub cmd_alterar_Click()
    AlterarRegistro
End Sub
Private Sub cmd_anterior_Click()
    If Grupo.LocalizarAnterior Then
        MostraDados
    End If
End Sub
Private Sub cmd_excluir_Click()
    If MsgBox("confirma EXCLUSÃO Deste registro ? ", vbYesNo, "EXCLUIR") = vbYes Then
        ExcluirRegistro
    End If
End Sub

Private Sub cmd_incluir_Click()
    LimpaTela
    txt_codigo.SetFocus
End Sub

Private Sub cmd_ok_Click()
    IncluiRegistro
End Sub
Private Sub cmd_primeiro_Click()
    If Grupo.LocalizarPrimeiro Then
        MostraDados
    End If
End Sub
Private Sub cmd_proximo_Click()
    If Grupo.LocalizarProximo Then
        MostraDados
    End If
End Sub
Private Sub cmd_ultimo_Click()
    If Grupo.LocalizarUltimo Then
        MostraDados
    End If
End Sub
Private Sub Command1_Click()
    Dim xSQL As String
    Dim rsGrupo As New ADODB.Recordset
    xSQL = "SELECT Codigo, Nome FROM Grupo ORDER BY Nome"
    Set rsGrupo = Conectar.RstConexao(xSQL)
    With rsGrupo
        If Not .EOF Then
            .MoveFirst
            Do Until .EOF
                MsgBox rsGrupo("Nome").Value & " - " & rsGrupo("Codigo").Value
                .MoveNext
            Loop
        Else
            MsgBox "Tabela não contém registro!"
        End If
    End With
    Set rsGrupo = Nothing
End Sub
Private Sub Form_Load()
    Set Grupo = New cGrupo
    Set Conectar = New CConexao
    Set Grupo.Conexao = Conectar.Conexao
    If Grupo.LocalizarUltimo Then
        MostraDados
    End If
End Sub
Private Sub MostraDados()
    txt_codigo.Text = Grupo.Codigo
    txt_nome.Text = Grupo.Nome
End Sub
Private Sub IncluiRegistro()
    Grupo.Codigo = "" & txt_codigo.Text
    Grupo.Nome = "" & txt_nome.Text

    If (Not Grupo.Incluir) Then
        MsgBox "Não foi possivel incluir grupo!", vbCritical, "Erro"
        Exit Sub
    End If
End Sub
Private Sub AlterarRegistro()
    Grupo.Nome = "" & txt_nome.Text

    If (Not Grupo.Alterar(txt_codigo.Text)) Then
        MsgBox "Não foi possivel ALTERAR dados do grupo!", vbCritical, "Erro"
        Exit Sub
    Else
        MsgBox " Registro : " & txt_codigo.Text & " ALTERADO com sucesso ! ", vbInformation, "ALTERAR"
    End If
End Sub
Private Sub ExcluirRegistro()
    If (Not Grupo.Excluir(txt_codigo.Text)) Then
        MsgBox "Não foi possivel Excluir os dados do grupo!", vbCritical, "EXCLUIR"
        Exit Sub
    Else
        MsgBox " Registro : " & txt_codigo.Text & " EXCLUIDO com sucesso ! ", vbInformation, "ALTERAR"
        Grupo.LocalizarUltimo
        MostraDados
    End If
End Sub
Private Sub LimpaTela()
    Dim i As Integer
    For i = 0 To frm_grupo.Controls.Count - 1
        If TypeOf frm_grupo.Controls(i) Is TextBox Then
            frm_grupo.Controls(i).Text = ""
        End If
    Next
End Sub



Private Sub Image1_Click()
On Error GoTo localiza_erro
Dim resposta As Integer

resposta = InputBox("Informe o Código do grupo a localizar", "Código Grupo", 1, 3500, 3000)
 If resposta <> 0 Then
   If Grupo.LocalizarCodigo(resposta) Then
      MostraDados
   Else
      MsgBox "Não foi possivel localizar o grupo!", vbCritical, "Localizar"
   End If
 End If
 Exit Sub

localiza_erro:
 If Err.Number = 13 Then
   Resume Next
 Else
   MsgBox Err.Number & " - " & Err.Description
 End If

End Sub
