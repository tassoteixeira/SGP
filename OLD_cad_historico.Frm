VERSION 5.00
Object = "{BE4F3AC8-AEC9-101A-947B-00DD010F7B46}#1.0#0"; "MSOUTL32.OCX"
Begin VB.Form cadastro_historico 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Histórico"
   ClientHeight    =   3705
   ClientLeft      =   2250
   ClientTop       =   1830
   ClientWidth     =   6015
   ControlBox      =   0   'False
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
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3705
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5100
      Picture         =   "cad_historico.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2760
      Width           =   795
   End
   Begin VB.TextBox txt_nome 
      DataField       =   "desscricao"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1470
      MaxLength       =   40
      TabIndex        =   0
      Top             =   90
      Width           =   4425
   End
   Begin MSOutl.Outline Outline1 
      Height          =   2235
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5775
      _Version        =   65536
      _ExtentX        =   10186
      _ExtentY        =   3942
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Conta:"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   90
      TabIndex        =   2
      Top             =   120
      Width           =   1395
   End
End
Attribute VB_Name = "cadastro_historico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagInsere As Integer
Dim lOpcao As Integer
Dim lCodigo As Long
Dim lPai As String
Dim lInclusaoOk As Integer
Dim lDebitoCredito As String
Dim lPaiOld As String
Private Historico As cHistorico
Private Sub AtualizaDados()
    On Error GoTo FileError
    'Dim i As Integer
    'Dim nivel_expandir As Integer
    Dim x_codigo As Long
    If lOpcao = 1 Then 'Incluir
        If txt_nome.Text <> "" Then
            xCodigo = Outline1.Indent(Outline1.ListIndex)
            Historico.Empresa = g_empresa
            Historico.Codigo = Historico.ProximoCodigo(g_empresa)
            Historico.Nivel = xCodigo
            Historico.NomePai = lPai
            Historico.NomeFilho = txt_nome.Text
            Historico.DebitoCredito = lDebitoCredito
            'nivel_expandir = (Outline1.Indent(Outline1.ListIndex))
            If Historico.Incluir Then
                lCodigo = xCodigo
                AtualizaOutline
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
            'For i = 1 To nivel_expandir
            '    Outline1.Expand(i) = True
            'Next
        Else
            MsgBox "Digite o nome do Histórico!", 16, "Ajuda"
            txt_nome.SetFocus
        End If
    ElseIf lOpcao = 2 Then 'alterar
        lOpcao = 0
        Historico.NomeFilho = txt_nome.Text
        If Not Historico.Alterar(g_empresa, lCodigo) Then
            MsgBox "Não foi alterar este registro!", vbInformation, "Erro de Verificação!"
        End If
        'Falta rotina para alterar todos os registros filhos
        'Colocando o nome do novo Pai
        AtualizaOutline
    ElseIf lOpcao = 3 Then 'Deletar
        lOpcao = 0
        'Falta Verificar se tem registros filhos
        'Caso tenha informar que não pode excluir
        If MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro.") = 6 Then
            If Historico.Excluir(g_empresa, lCodigo) Then
                AtualizaOutline
            Else
                MsgBox "Não foi possivel excluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_historico.Name, "Históricoo"
    Exit Sub
End Sub
Private Sub AtualizaOutline()
    Dim arq_aux As Table
    Dim i As Integer
    Dim i2 As Integer
    Dim atualizou As Integer
    Dim posicao As Integer
    Dim ind As Integer
    Dim nivel_old As Integer
    txt_nome = ""
    Set arq_aux = tbl_historico.Clone()
    Outline1.Clear
    Outline1.AddItem "\"
    Outline1.Indent(0) = 1
    'ind = 0
    arq_aux.Index = "id_ordem"
    arq_aux.Seek ">", g_empresa, 0, " ", " "
    If Not arq_aux.NoMatch Then
        'monta somente contas do nível 1
        Do While Not arq_aux.EOF
            If arq_aux!Empresa <> g_empresa Then
                Exit Do
            End If
            If arq_aux!Nivel > 1 Then
                Exit Do
            End If
            ind = ind + 1
            Outline1.AddItem Trim(arq_aux![Nome Filho])
            Outline1.Indent(ind) = arq_aux!Nivel + 1
            arq_aux.MoveNext
            If arq_aux.EOF Then
                Exit Do
            End If
        Loop
        'mostra aberto dados do registro 0
        Outline1.Expand(0) = True
    End If
    nivel_old = 2
    If Not (arq_aux.EOF And arq_aux.BOF) Then
        If Not arq_aux.NoMatch Then
            Do While Not arq_aux.EOF
                If arq_aux!Empresa <> g_empresa Then
                    Exit Do
                End If
                'percorre outline1 até o fim
                For i = 0 To Outline1.ListCount - 1
                    'testa se o nível_outline1 encontrado é igual ao arq_aux!Nível e
                    'se o nome_outline1 é igual ao arq_aux![Nome Filho]
                    If Outline1.Indent(i) = arq_aux!Nivel And Outline1.List(i) = Trim(arq_aux![Nome Pai]) Then
                        atualizou = 0
                        For i2 = i To Outline1.ListCount - 1
                            If Outline1.Indent(i2) = arq_aux!Nivel And Outline1.List(i2) = Trim(arq_aux![Nome Pai]) Then
                                posicao = i2 + 1
                            End If
                            If Outline1.Indent(i2) = arq_aux!Nivel And Outline1.List(i2) = Trim(arq_aux![Nome Pai]) And Outline1.List(i2) < Trim(arq_aux![Nome Pai]) Then
                                Outline1.AddItem Trim(arq_aux![Nome Filho]), i2
                                Outline1.Indent(i2) = arq_aux!Nivel + 1
                                atualizou = 1
                                Exit For
                            End If
                        Next
                        If atualizou = 0 Then
                            For i2 = posicao To Outline1.ListCount - 1
                                If Outline1.Indent(i2) <> (arq_aux!Nivel + 1) Then
                                    Exit For
                                End If
                                posicao = posicao + 1
                            Next
                            Outline1.AddItem Trim(arq_aux![Nome Filho]), posicao
                            Outline1.Indent(posicao) = arq_aux!Nivel + 1
                            Exit For
                        End If
                    End If
                Next
                arq_aux.MoveNext
                If Not arq_aux.EOF Then
                    If arq_aux!Empresa <> g_empresa Then
                        Exit Do
                    End If
                End If
            Loop
        End If
    End If
    If Outline1.ListCount < 3 Then
        lPai = "\"
    End If
End Sub
Private Sub AtualTabe()
    tbl_historico!Codigo = lCodigo
    tbl_historico!descricao = txt_nome
End Sub
Private Sub TestaPai()
    If lPai = "\" Then
        tbl_historico.Index = "id_pai"
        tbl_historico.Seek ">=", g_empresa, lPai, ""
        If Not tbl_historico.NoMatch Then
            If tbl_historico!Empresa = g_empresa Then
                lDebitoCredito = "C"
                lInclusaoOk = True
                Exit Sub
            Else
                lDebitoCredito = "D"
                lInclusaoOk = True
                Exit Sub
            End If
        Else
            lDebitoCredito = "D"
            lInclusaoOk = True
            Exit Sub
        End If
        If tbl_historico![Nome Pai] <> "\" Then
            lInclusaoOk = True
            Exit Sub
        End If
        tbl_historico.MoveNext
        If Not tbl_historico.EOF Then
            If tbl_historico!Empresa = g_empresa Then
                If tbl_historico![Nome Pai] <> "\" Then
                    lInclusaoOk = True
                    Exit Sub
                Else
                    lInclusaoOk = False
                End If
            End If
        Else
            lInclusaoOk = True
        End If
    Else
        lInclusaoOk = True
    End If
End Sub
Private Sub TestaEmpresa()
    If g_empresa > 0 Then
        AtualizaOutline
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Finaliza()
  If txt_nome = "" Then
    If (MsgBox("Não foi escolhido nenhum histórico!", 1, "Erro de Verificação")) = 1 Then
        g_historico = 0
    End If
  End If
End Sub
Private Sub Form_Activate()
    TestaEmpresa
End Sub
Private Sub Form_Load()
    CentraForm Me
    Screen.MousePointer = 1
    If Not Screen.ActiveForm Is movimento_bancario Then
        Outline1.ForeColor = &HFF&
    End If
    Set tbl_historico = bd_sgp.OpenTable("Historico")
    TestaEmpresa
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub Outline1_Click()
    Dim x_pai As String
    Dim x_filho As String
    lFlagInsere = False
    If Outline1.Text <> "\" Then
        x_pai = Outline1.FullPath(Outline1.ListIndex)
        x_filho = ""
        Call PaiFilho(x_pai, x_filho)
        'tbl_historico.Index = "id_pai"
        If Outline1.Indent(Outline1.ListIndex) > 2 Then
            'tbl_historico.Seek "=", g_empresa, x_pai, x_filho
            'If Not tbl_historico.NoMatch Then
            If Historico.LocalizarPai(g_empresa, x_pai, x_filho) Then
                lPaiOld = Historico.NomeFilho
                txt_nome.Text = Historico.NomeFilho
                lDebitoCredito = Historico.DebitoCredito
                g_string = Historico.Codigo & "|@|"
                'lFlagInsere = True
            Else
                MsgBox "Item não encontrado!", 16, "Erro de Verificação!"
                txt_nome.Text = ""
                g_string = 0 & "|@|"
                'lFlagInsere = False
            End If
        End If
    Else
        lPai = Outline1.Text
        'TestaPai
        'If lInclusaoOk Then
        '    lPai = "\"
            txt_nome = lPai
        '    'lFlagInsere = True
        'Else
        '    MsgBox "Escolha outra conta para cadastro!", 16, "Ajuda"
        '    Outline1.SetFocus
        'End If
    End If
End Sub
Private Sub Outline1_DblClick()
    Unload Me
End Sub
Private Sub Outline1_KeyDown(KeyCode As Integer, Shift As Integer)
    lFlagInsere = False
    If KeyCode = 45 Then
        lFlagInsere = True
        txt_nome.Text = ""
        lOpcao = 1
        Dim auxlCodigo As Long
        Dim guardalCodigo As String
        Dim tam As Long
        lCodigo = Outline1.Indent(Outline1.ListIndex)
        lPai = Outline1.List(Outline1.ListIndex)
        txt_nome.SetFocus
        'MsgBox "Escolha um item para inserir!", 16, "Erro de Verificação!"
    ElseIf KeyCode = 46 Then
        If Outline1.Expand(Outline1.ListIndex) = False Then
            lOpcao = 3
            AtualizaDados
        Else
            MsgBox "Para apagar este item você tem que apagar os Subitens!", 16, "Erro de Verificação!"
        End If
    ElseIf KeyCode = 13 Then
        AtualizaDados
    End If
End Sub
Private Sub Outline1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Finaliza
    End If
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If lFlagInsere Then
            AtualizaDados
            Outline1.SetFocus
        Else
            lOpcao = 2
            AtualizaDados
            Outline1.SetFocus
        End If
    End If
End Sub
