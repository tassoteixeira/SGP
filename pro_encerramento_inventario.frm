VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form pro_encerramento_inventario 
   Caption         =   "Encerramento do Inventário"
   ClientHeight    =   2550
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   7575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2550
   ScaleWidth      =   7575
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   2640
      Picture         =   "pro_encerramento_inventario.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Confirma o encerramento do inventário."
      Top             =   1620
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4380
      Picture         =   "pro_encerramento_inventario.frx":160A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1620
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1455
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7455
      Begin VB.ComboBox cboProduto 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1020
         Width           =   5475
      End
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   5475
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   285
         Left            =   1860
         TabIndex        =   2
         Top             =   240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "P&roduto"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1020
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "&Grupo"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "D&ata do Encerramento"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "pro_encerramento_inventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Estoque As New cEstoque
Private Estoque2 As New cEstoque2
Private EntradaCombustivel As New cEntradaCombustivel
Private MedicaoCombustivel As New cMedicaoCombustivel
Private MovimentoBomba As New cMovimentoBomba
Dim rs As New adodb.Recordset
Dim lSQL As String

Private Sub GeraEncerramentoInventario()
    Dim rsProduto As adodb.Recordset
    Dim xSQL As String
    Dim xString As String
    
    Set rsProduto = New adodb.Recordset
    xString = " WHERE"
    xSQL = ""
    xSQL = xSQL & "SELECT Codigo, [Preco de Custo], [Preco de Custo Medio], [Codigo do Grupo], [Tipo de Combustivel]"
    xSQL = xSQL & "  FROM Produto"
    If cboGrupo.ItemData(cboGrupo.ListIndex) > 0 Then
        xSQL = xSQL & xString & " [Codigo do Grupo] = " & cboGrupo.ItemData(cboGrupo.ListIndex)
        xString = " AND"
    End If
    If cboProduto.ItemData(cboProduto.ListIndex) > 0 Then
        xSQL = xSQL & xString & " Codigo = " & cboProduto.ItemData(cboProduto.ListIndex)
        xString = " AND"
    End If
    xSQL = xSQL & xString & " [Tipo de Produto] NOT IN (7,8) "
    
    xSQL = xSQL & " ORDER BY [Codigo do Grupo], Nome, Codigo"
    Set rsProduto = Conectar.RsConexao(xSQL)
    If rsProduto.RecordCount > 0 Then
        Do Until rsProduto.EOF
            If Estoque.LocalizarCodigo(g_empresa, rsProduto("Codigo").Value) Then
                Estoque2.Empresa = g_empresa
                Estoque2.Data = CDate(msk_data.Text)
                Estoque2.GrupoProduto = Estoque.GrupoProduto
                Estoque2.CodigoProduto2 = Estoque.CodigoProduto2
                Estoque2.Quantidade = Estoque.Quantidade
                Estoque2.PrecoVenda = Estoque.PrecoVenda
                Estoque2.PrecoCusto = rsProduto("Preco de Custo").Value
                Estoque2.PrecoCustoMedio = rsProduto("Preco de Custo Medio").Value
                If Trim(rsProduto("Tipo de Combustivel").Value) <> "" Then
                    If MovimentoBomba.LocalizarPrimeiroBicoComb(g_empresa, CDate(msk_data.Text), rsProduto("Tipo de Combustivel").Value) Then
                        Estoque2.PrecoCusto = MovimentoBomba.PrecoCusto
                        Estoque2.PrecoVenda = MovimentoBomba.PrecoVenda
                    End If
                    Estoque2.Quantidade = MedicaoCombustivel.TotalMedidaCombustivel(g_empresa, CDate(msk_data.Text) + 1, rsProduto("Tipo de Combustivel").Value, 0)
                End If
                If Not Estoque2.Incluir Then
                    MsgBox "Não foi possível incluir o estoque2!", vbInformation, "Erro de Integridade!"
                End If
            End If
            rsProduto.MoveNext
        Loop
    End If
    rsProduto.Close
    Set rsProduto = Nothing
    MsgBox "Inventário Contábil Arquivado com sucesso!", vbInformation + vbOKOnly, "Processamento Concluído!"
End Sub
Private Sub PreencheCboGrupo()
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Codigo, Nome"
    lSQL = lSQL & "     FROM Grupo"
    lSQL = lSQL & " ORDER BY Nome, Codigo"
    'Abre RecordSet
    Set rs = New adodb.Recordset
    Set rs = Conectar.RsConexao(lSQL)
    
    cboGrupo.Clear
    cboGrupo.AddItem "Todos os Grupos"
    cboGrupo.ItemData(cboGrupo.NewIndex) = 0
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do Until rs.EOF
            cboGrupo.AddItem rs("Nome").Value
            cboGrupo.ItemData(cboGrupo.NewIndex) = rs("Codigo").Value
            rs.MoveNext
        Loop
    End If
End Sub
Private Sub PreencheCboProduto()
    cboProduto.Clear
    
    cboProduto.AddItem "Todos os Produtos"
    cboProduto.ItemData(cboProduto.NewIndex) = 0
    lSQL = ""
    lSQL = lSQL & "SELECT Codigo, Nome"
    lSQL = lSQL & "  FROM Produto"
    lSQL = lSQL & " ORDER BY Nome, Codigo"
    Set rs = Conectar.RsConexao(lSQL)
    If rs.RecordCount > 0 Then
        Do Until rs.EOF
                cboProduto.AddItem rs!Nome
                cboProduto.ItemData(cboProduto.NewIndex) = rs!Codigo
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Estoque = Nothing
    Set Estoque2 = Nothing
    Set EntradaCombustivel = Nothing
    Set MedicaoCombustivel = Nothing
    Set MovimentoBomba = Nothing
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data do encerramento do inventário.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf cboGrupo.ListIndex = -1 Then
        MsgBox "Selecione o grupo.", vbInformation, "Atenção!"
        cboGrupo.SetFocus
    ElseIf cboProduto.ListIndex = -1 Then
        MsgBox "Selecione o produto.", vbInformation, "Atenção!"
        cboProduto.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cboGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboProduto.SetFocus
    End If
End Sub
Private Sub cboProduto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub cmd_ok_Click()
    If ValidaCampos Then
        If Estoque2.ExisteEstoqueData(g_empresa, CDate(msk_data.Text), cboGrupo.ItemData(cboGrupo.ListIndex), cboProduto.ItemData(cboProduto.ListIndex)) = False Then
            If (MsgBox("Deseja realmente gerar o encerramento do inventário?", vbYesNo + vbDefaultButton2, "Encerramento do Inventário!")) = 6 Then
                Call GravaAuditoria(1, Me.name, 26, "Gera Encerramento Inventário Empresa:" & g_empresa & "-" & Mid(g_nome_empresa, 1, 20) & " Em:" & msk_data.Text)
                cmd_ok.Enabled = False
                cmd_sair.Enabled = False
                GeraEncerramentoInventario
                cmd_ok.Enabled = True
                cmd_sair.Enabled = True
            End If
        Else
            MsgBox "Inventário já encerrado nesta data!", vbInformation, "Atenção!"
            If (MsgBox("Deseja deletar o encerramento do inventário desta data?", vbQuestion + vbYesNo + vbDefaultButton2, "Deleta Encerramento do Inventário")) = 6 Then
                Call GravaAuditoria(1, Me.name, 26, "Deleta Encerramento Inventário Empresa:" & g_empresa & "-" & Mid(g_nome_empresa, 1, 20) & " Em:" & msk_data.Text)
                If Not Estoque2.ExcluirData(g_empresa, CDate(msk_data.Text), cboGrupo.ItemData(cboGrupo.ListIndex), cboProduto.ItemData(cboProduto.ListIndex)) Then
                    MsgBox "Não foi possível excluir o estoque2!", vbInformation, "Erro de Integridade!"
                End If
            End If
        End If
        cmd_sair.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(Date, "dd/MM/yyyy")
        cboGrupo.ListIndex = 0
        cboProduto.ListIndex = 0
        msk_data.SetFocus
    End If
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
'    If g_nome_usuario = "L.M.C." Then
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel_LMC"
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivelLMC"
        MovimentoBomba.NomeTabela = "Movimento_Bomba_LMC"
'    Else
'        MedicaoCombustivel.NomeTabela = "MedicaoCombustivel"
'    End If
    PreencheCboGrupo
    PreencheCboProduto
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_GotFocus()
    msk_data.SelStart = 0
    msk_data.SelLength = 5
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboGrupo.SetFocus
    End If
End Sub
