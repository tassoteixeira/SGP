VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form processamento_custo_produto 
   Caption         =   "Processamento de Custo de Produtos"
   ClientHeight    =   3390
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   6495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "processamento_custo_produto.frx":0000
   ScaleHeight     =   3390
   ScaleWidth      =   6495
   Begin VB.Frame frmDados 
      Height          =   2235
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6255
      Begin VB.OptionButton optCustoRealMedio 
         Caption         =   "Custo Real Médio"
         Height          =   255
         Left            =   300
         TabIndex        =   2
         Top             =   660
         Width           =   2595
      End
      Begin VB.OptionButton optCustoMedio 
         Caption         =   "Custo Médio"
         Height          =   255
         Left            =   300
         TabIndex        =   3
         Top             =   1020
         Width           =   2595
      End
      Begin VB.OptionButton optCustoReal 
         Caption         =   "Custo Real"
         Height          =   255
         Left            =   300
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   2595
      End
      Begin MSMask.MaskEdBox msk_data_inicial 
         Height          =   300
         Left            =   180
         TabIndex        =   5
         Top             =   1740
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
         TabIndex        =   7
         Top             =   1740
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
         TabIndex        =   6
         Top             =   1530
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Data inicial"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   4
         Top             =   1530
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   1680
      Picture         =   "processamento_custo_produto.frx":0446
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Confirma o processamento."
      Top             =   2460
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4020
      Picture         =   "processamento_custo_produto.frx":1A50
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2460
      Width           =   795
   End
End
Attribute VB_Name = "processamento_custo_produto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lSQL As String
Private rsMovimentoLubrificante As New adodb.Recordset

Private EntradaProduto As New cEntradaProduto
Private MovimentoLubrificante As New cMovimentoLubrificante
Private Produto As New cProduto
Private Sub Finaliza()
    Set EntradaProduto = Nothing
    Set MovimentoLubrificante = Nothing
    Set Produto = Nothing
End Sub
Private Sub Processamento()
    If optCustoReal.Value = True Then
        ProcessamentoCustoReal
    End If
End Sub
Private Sub ProcessamentoCustoReal()
    Dim xData As Date
    Dim xPreco As Currency
    Dim xValor(0 To 10) As Currency
    Dim xQuantidade(0 To 10)  As Currency
    Dim xQtdVendaDia As Currency
    Dim xQtdEntradaDia As Currency
    Dim xQtdDeQuantidade As Integer
    Dim xTipoCombustivel As String
    Dim xString As String
    Dim xQtdCusto As Currency
    Dim xValorCusto As Currency
    Dim i As Integer
    
    
    lSQL = ""
    lSQL = lSQL & "SELECT Data, [Codigo do Produto2]"
    lSQL = lSQL & "  FROM Movimento_Lubrificante"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data >= " & preparaData(CDate(msk_data_inicial.Text))
    lSQL = lSQL & "   AND Data <= " & preparaData(CDate(msk_data_final.Text))
    lSQL = lSQL & " ORDER BY Data, [Codigo do Produto2]"
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será calculado o custo real de combustível entre " & msk_data_inicial.Text & " a " & msk_data_final.Text & "." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Calcula Custo Real de Produtos!")) = vbYes Then
        Set rsMovimentoLubrificante = Conectar.RsConexao(lSQL)
        If rsMovimentoLubrificante.RecordCount > 0 Then
            Do Until rsMovimentoLubrificante.EOF
                If Produto.LocalizarCodigo(rsMovimentoLubrificante![Codigo do Produto2]) Then
                    If Not MovimentoLubrificante.AlteraCusto(g_empresa, rsMovimentoLubrificante!Data, rsMovimentoLubrificante![Codigo do Produto2], Produto.PrecoCusto) Then
                        MsgBox "Não foi possível alterar preço de custo!" & Chr(10) & "Codigo do Produto:" & rsMovimentoLubrificante![Codigo do Produto2], vbInformation, "Erro de Integridade!"
                    End If
                Else
                    MsgBox "Produto Não Cadastrado!" & Chr(10) & "Codigo:" & rsMovimentoLubrificante![Codigo do Produto2], vbInformation, "Erro de Integridade!"
                End If
                rsMovimentoLubrificante.MoveNext
            Loop
        End If
        rsMovimentoLubrificante.Close
        Set rsMovimentoLubrificante = Nothing
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com o custo real calculado.", vbInformation, "Processamento Concluído!"
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
    'ErroArquivo tbl_estoque.Name, "Estoqueo"
    Exit Sub
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data_inicial.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_inicial.SetFocus
    ElseIf Not IsDate(msk_data_final.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_final.SetFocus
    ElseIf Not IsDate(msk_data_final.Text) >= IsDate(msk_data_inicial.Text) Then
        MsgBox "A data final deve ser igual ou maior que " & msk_data_inicial & " .", vbInformation, "Atenção!"
        msk_data_final.SetFocus
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
    
    msk_data_inicial.Text = Format(g_data_def, "dd/mm/yyyy")
    msk_data_final.Text = Format(g_data_def, "dd/mm/yyyy")
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
Private Sub optCustoMedio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_inicial.SetFocus
    End If
End Sub
Private Sub optCustoReal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_inicial.SetFocus
    End If
End Sub
Private Sub optCustoRealMedio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_inicial.SetFocus
    End If
End Sub
