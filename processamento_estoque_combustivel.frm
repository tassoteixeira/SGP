VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form processamento_estoque_combustivel 
   Caption         =   "Processamento de Estoque de Combustível"
   ClientHeight    =   3615
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   6495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "processamento_estoque_combustivel.frx":0000
   ScaleHeight     =   3615
   ScaleWidth      =   6495
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   1680
      Picture         =   "processamento_estoque_combustivel.frx":0446
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Confirma o processamento."
      Top             =   2640
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4020
      Picture         =   "processamento_estoque_combustivel.frx":1A50
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2640
      Width           =   795
   End
   Begin VB.Frame frmDados
      Height          =   2475
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6255
      Begin VB.CheckBox chk_calcula_saida 
         Caption         =   "Ca&lcula as Saidas do Período"
         Height          =   300
         Left            =   180
         TabIndex        =   4
         Top             =   1320
         Width           =   5595
      End
      Begin VB.CheckBox chk_calcula_entrada 
         Caption         =   "&Calcula as Entradas do Período"
         Height          =   300
         Left            =   180
         TabIndex        =   3
         Top             =   960
         Width           =   5595
      End
      Begin VB.CheckBox chk_zera_estoque 
         Caption         =   "&Zera o Estoque Atual"
         Height          =   300
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   5595
      End
      Begin VB.CheckBox chk_move_entrada_estoque 
         Caption         =   "&Move as Entradas (Inventário) para o Estoque Atual"
         Height          =   300
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   5595
      End
      Begin MSMask.MaskEdBox msk_data_inicial 
         Height          =   300
         Left            =   180
         TabIndex        =   6
         Top             =   1980
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
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
      Begin MSMask.MaskEdBox msk_data_final 
         Height          =   300
         Left            =   2820
         TabIndex        =   8
         Top             =   1980
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
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
      Begin VB.Label Label3 
         Caption         =   "D&ata final"
         Height          =   195
         Index           =   8
         Left            =   2820
         TabIndex        =   7
         Top             =   1770
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Data inicial"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   5
         Top             =   1770
         Width           =   1815
      End
   End
End
Attribute VB_Name = "processamento_estoque_combustivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbl_entrada_combustivel As Table
Dim tbl_combustivel As Table
Dim tbl_movimento_bomba As Table
Private Sub Finaliza()
    tbl_entrada_combustivel.Close
    tbl_combustivel.Close
    tbl_movimento_bomba.Close
End Sub
Private Sub Processamento()
    If chk_zera_estoque.Value = 1 Then
        ProcessamentoZeraEstoque
    End If
    If chk_move_entrada_estoque.Value = 1 Then
        ProcessamentoMoveEntradaEstoque
    End If
    If chk_calcula_entrada.Value = 1 Then
        ProcessamentoCalculaEntradaEstoque
    End If
    If chk_calcula_saida.Value = 1 Then
        ProcessamentoCalculaSaidaEstoque
    End If
End Sub
Private Sub ProcessamentoCalculaEntradaEstoque()
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será calculado as entradas de " & msk_data_inicial & " até " & msk_data_final & " para todo o estoque de combustíveis." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Calcula Entradas para o Estoque de Combustíveis!")) = 6 Then
        With tbl_entrada_combustivel
            .Seek ">", g_empresa, CDate(msk_data_inicial), " ", " "
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Or !Data > CDate(msk_data_final) Then
                        Exit Do
                    End If
                    If ![Numero da Nota] <> 1 Then
                        tbl_combustivel.Seek "=", g_empresa, ![Tipo de Combustivel]
                        If Not tbl_combustivel.NoMatch Then
                            tbl_combustivel.Edit
                            tbl_combustivel![Quantidade em Estoque] = tbl_combustivel![Quantidade em Estoque] + !Quantidade
                            tbl_combustivel.Update
                        End If
                    End If
                    .MoveNext
                Loop
            End If
        End With
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com as entradas calculadas para o estoque de combustíveis.", vbInformation, "Operação Concluida!"
    End If
End Sub
Private Sub ProcessamentoCalculaSaidaEstoque()
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será calculado as saidas de combustíveis de " & msk_data_inicial & " até " & msk_data_final & " para todo estoque." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Calcula Saidas para o Estoque de Combustíveis!")) = 6 Then
        With tbl_movimento_bomba
            .Index = "id_data"
            .Seek ">", g_empresa, CDate(msk_data_inicial), 0, 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Or !Data > CDate(msk_data_final) Then
                        Exit Do
                    End If
                    tbl_combustivel.Seek "=", g_empresa, ![Tipo de Combustivel]
                    If Not tbl_combustivel.NoMatch Then
                        tbl_combustivel.Edit
                        tbl_combustivel![Quantidade em Estoque] = tbl_combustivel![Quantidade em Estoque] - ![Quantidade da Saida]
                        tbl_combustivel.Update
                    End If
                    .MoveNext
                Loop
            End If
        End With
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com as saidas calculadas para o estoque.", vbInformation, "Operação Concluida!"
    End If
End Sub
Private Sub ProcessamentoMoveEntradaEstoque()
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será movido as entradas de combustível (Nota = 1) da data " & msk_data_inicial & " para todo estoque." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Move Entradas para o Estoque!")) = 6 Then
        tbl_combustivel.Seek ">=", g_empresa, "  "
        If Not tbl_combustivel.NoMatch Then
            Do Until tbl_combustivel.EOF
                If tbl_combustivel!Empresa <> g_empresa Then
                    Exit Do
                End If
                With tbl_entrada_combustivel
                .Seek "=", g_empresa, CDate(msk_data_inicial), tbl_combustivel!Codigo, 1
                    If Not .NoMatch Then
                        tbl_combustivel.Edit
                        tbl_combustivel![Quantidade em Estoque] = !Quantidade
                        tbl_combustivel.Update
                    End If
                End With
                tbl_combustivel.MoveNext
            Loop
        End If
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com as entradas (inventário) movidas para o estoque.", vbInformation, "Operação Concluida!"
    End If
End Sub
Private Sub ProcessamentoZeraEstoque()
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será feito o processamento para zerar todo seu estoque de combustível." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Zeramento de Estoque de Combustível!")) = 6 Then
        bd_sgp.Execute "Update Combustivel Set [Quantidade em Estoque] = 0 Where Empresa = " & g_empresa
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com todo seu estoque de combustível zerado.", vbInformation, "Operação Concluida!"
    End If
End Sub
Private Sub chk_calcula_entrada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_calcula_saida.SetFocus
    End If
End Sub
Private Sub chk_calcula_saida_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_inicial.SetFocus
    End If
End Sub
Private Sub chk_move_entrada_estoque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_calcula_entrada.SetFocus
    End If
End Sub
Private Sub chk_zera_estoque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_move_entrada_estoque.SetFocus
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        Processamento
        cmd_sair.SetFocus
    End If
    Exit Sub
FileError:
    ErroArquivo tbl_combustivel.name, "Combustívelo"
    Exit Sub
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data_inicial) Then
        MsgBox "Informe a data inicial.", 64, "Atenção!"
        msk_data_inicial.SetFocus
    ElseIf Not IsDate(msk_data_final) Then
        MsgBox "Informe a data final.", 64, "Atenção!"
        msk_data_final.SetFocus
    ElseIf Not IsDate(msk_data_final) >= IsDate(msk_data_inicial) Then
        MsgBox "A data final deve ser igual ou maior que " & msk_data_inicial & " .", 64, "Atenção!"
        msk_data_final.SetFocus
    ElseIf chk_zera_estoque.Value = 0 And chk_move_entrada_estoque.Value = 0 And chk_calcula_entrada.Value = 0 And chk_calcula_saida.Value = 0 Then
        MsgBox "Deve ser selecionada pelo menos uma das opções acima.", 64, "Atenção!"
        chk_zera_estoque.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Screen.MousePointer = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_sair_Click
    ElseIf KeyCode = vbKeyF11 Then
        KeyCode = 0
        cmd_ok_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    Set tbl_entrada_combustivel = bd_sgp.OpenTable("Entrada_Combustivel")
    Set tbl_combustivel = bd_sgp.OpenTable("Combustivel")
    Set tbl_movimento_bomba = bd_sgp.OpenTable("Movimento_Bomba")
    tbl_entrada_combustivel.Index = "id_data"
    tbl_combustivel.Index = "id_codigo"
    tbl_movimento_bomba.Index = "id_data"
    msk_data_inicial = g_data_def
    msk_data_final = g_data_def
    chk_zera_estoque.Value = 0
    chk_move_entrada_estoque.Value = 1
    chk_calcula_entrada.Value = 1
    chk_calcula_saida.Value = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_final_GotFocus()
    msk_data_final.SelStart = 0
    msk_data_final.SelLength = 5
End Sub
Private Sub msk_data_final_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub msk_data_inicial_GotFocus()
    msk_data_inicial.SelStart = 0
    msk_data_inicial.SelLength = 5
End Sub
Private Sub msk_data_inicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_final.SetFocus
    End If
End Sub
