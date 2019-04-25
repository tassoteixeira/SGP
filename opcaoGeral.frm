VERSION 5.00
Begin VB.Form opcaoGeral 
   Caption         =   "Escolha a Opção"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "opcaoGeral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.ComboBox cboOpcao 
      Height          =   315
      Left            =   660
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   660
      Width           =   3615
   End
   Begin VB.Label lblObservacao 
      Caption         =   "Observação"
      Height          =   1155
      Left            =   660
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Escolha a Opção Desejada"
      Height          =   255
      Left            =   660
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "opcaoGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboOpcao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Unload Me
    End If
End Sub
Private Sub Finaliza()
    g_string = ""
    If cboOpcao.ListIndex <> -1 Then
        g_string = g_string & cboOpcao.ItemData(cboOpcao.ListIndex) & "|@|"
        g_string = g_string & cboOpcao.Text & "|@|"
    End If
End Sub
Private Sub PreencheCboOpcao()
    Dim xQtdOpcao As Integer
    Dim i As Integer
    Dim i2 As Integer
    
    cboOpcao.Clear
    If Len(g_string) > 0 Then
        xQtdOpcao = RetiraGString(2)
        For i = 1 To xQtdOpcao
            i2 = i * 2 + 2
            cboOpcao.AddItem RetiraGString(i2)
            i2 = i * 2 + 1
            If Len(RetiraGString(i2)) > 6 Then
                cboOpcao.ItemData(cboOpcao.NewIndex) = i2 - 2
            Else
                cboOpcao.ItemData(cboOpcao.NewIndex) = RetiraGString(i2)
            End If
        Next
    End If
    If xQtdOpcao > 2 Then
        cboOpcao.ListIndex = 2
    Else
        cboOpcao.ListIndex = 0
    End If
End Sub
Private Sub Form_Activate()
    Screen.MousePointer = 1
End Sub
Private Sub Form_Load()
    CentraForm Me
    PreencheCboOpcao
    If Len(g_string) > 0 Then
        lblObservacao.Caption = RetiraGString(1)
    End If
    If lblObservacao.Caption = "Selecione a Empresa para Conexão!" Then
        Me.Caption = "Conexão Multi-Empresa"
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
