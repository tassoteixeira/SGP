VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ConsultaProdutoCodigoBarra 
   Caption         =   "Consulta de Produtos por Código de Barra"
   ClientHeight    =   5355
   ClientLeft      =   2130
   ClientTop       =   1305
   ClientWidth     =   8655
   Icon            =   "ConsultaProdutoCodigoBarra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "ConsultaProdutoCodigoBarra.frx":0442
   ScaleHeight     =   5355
   ScaleWidth      =   8655
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6855
      Begin VB.TextBox txt_condicao 
         Height          =   285
         Left            =   1980
         TabIndex        =   2
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "Código de &Barra"
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   480
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   7020
      Picture         =   "ConsultaProdutoCodigoBarra.frx":0488
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Inicia a pesquisa selecionada."
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   7860
      Picture         =   "ConsultaProdutoCodigoBarra.frx":1A92
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   120
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   4035
      Left            =   60
      TabIndex        =   5
      Top             =   1260
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7117
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   12582912
      ForeColorSel    =   16777215
      BackColorBkg    =   12632256
      SelectionMode   =   1
   End
End
Attribute VB_Name = "ConsultaProdutoCodigoBarra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim lFlagConsulta As Integer
    Dim lSQL As String
    
    Dim fld As adodb.Field
    Dim lRst As adodb.Recordset
Private Sub AtualizaGrid()
    On Error GoTo ErroConsulta
    
    If ValidaCampos Then
        'Prepara SQL
        lSQL = ""
        lSQL = lSQL & "SELECT Produto.Codigo, Produto.Nome, Produto.Unidade, Produto.[Preco de Custo],"
        lSQL = lSQL & "       Produto.[Margem de Lucro Desejada], Produto.[Preco de Venda], Estoque.Quantidade,"
        lSQL = lSQL & "       Grupo.Nome As NomeGrupo, Sub_Grupo.Nome As NomeSubGrupo, Produto.Comissao,"
        lSQL = lSQL & "       Produto.[Codigo de Barra]"
        lSQL = lSQL & "  FROM Produto, Estoque, Grupo, Sub_Grupo"
        lSQL = lSQL & " WHERE Grupo.Codigo = Produto.[Codigo do Grupo]"
        lSQL = lSQL & "   AND Sub_Grupo.Codigo = Produto.[Codigo do SubGrupo]"
        lSQL = lSQL & "   AND Estoque.Empresa = " & g_empresa
        lSQL = lSQL & "   AND Estoque.[Codigo do Produto2] = Produto.Codigo"
        lSQL = lSQL & " AND Produto.[Codigo de Barra] = " & preparaTexto(txt_condicao.Text)
        'If lOcultar Then
            lSQL = lSQL & "   AND Produto.Inativo = " & preparaBooleano(False)
        'End If
        lSQL = lSQL & " ORDER BY Produto.Nome"
        'Abre RecordSet
        Set lRst = Nothing
        Set lRst = New adodb.Recordset
        Set lRst = Conectar.RsConexao(lSQL)
        FormaGrid
        MontaGrid
    End If
    Exit Sub
ErroConsulta:
    If Err = 3075 Then
        MsgBox "Condição inválida.", vbInformation, "Erro de Consulta"
    Else
        MsgBox Error, vbInformation, "Erro de Consulta"
    End If
    Exit Sub
End Sub
Private Sub Finaliza()
    Set lRst = Nothing
End Sub
Private Sub FormaGrid()
    Dim i As Integer
    
    MSFlexGrid.WordWrap = True
    MSFlexGrid.Cols = 11
    MSFlexGrid.RowHeight(0) = 650
    MSFlexGrid.Rows = 1
    i = 0
    MSFlexGrid.TextMatrix(0, i) = "Código"
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 600
    i = i + 1
    MSFlexGrid.TextMatrix(0, i) = "Nome do Produto"
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 3600
    i = i + 1
    MSFlexGrid.TextMatrix(0, i) = "Un."
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 450
    i = i + 1
    MSFlexGrid.TextMatrix(0, i) = "Preço de Custo"
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 1000
    'If lOcultar Then
        MSFlexGrid.ColWidth(i) = 1
    'End If
    i = i + 1
    MSFlexGrid.TextMatrix(0, i) = "Mardem de Lucro Desejada"
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 1000
    'If lOcultar Then
        MSFlexGrid.ColWidth(i) = 1
    'End If
    i = i + 1
    MSFlexGrid.TextMatrix(0, i) = "Preço de Venda"
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 1000
    i = i + 1
    MSFlexGrid.TextMatrix(0, i) = "Quantidade"
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 950
    'If lOcultar Then
        MSFlexGrid.ColWidth(i) = 1
    'End If
    i = i + 1
    MSFlexGrid.TextMatrix(0, i) = "Grupo"
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 2000
    i = i + 1
    MSFlexGrid.TextMatrix(0, i) = "SubGrupo"
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 2000
    i = i + 1
    MSFlexGrid.TextMatrix(0, i) = "Comissão"
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 900
    'If lOcultar Then
        MSFlexGrid.ColWidth(i) = 1
    'End If
    i = i + 1
    MSFlexGrid.TextMatrix(0, i) = "Código de Barra"
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 1300
End Sub
Private Sub MarcaCelulas()
    g_string = ""
    If Val(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0)) > 0 Then
        g_string = MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0) & "|@|"
    End If
End Sub
Private Sub MontaGrid()
    Dim i As Integer
    Dim i2 As Integer
    
    If Not lRst.EOF Then
        MSFlexGrid.Visible = False
        lRst.MoveFirst
        Do Until lRst.EOF
            MSFlexGrid.Rows = MSFlexGrid.Rows + 1
            MSFlexGrid.Row = MSFlexGrid.Rows - 1
            i2 = -1
            For Each fld In lRst.Fields
                i2 = i2 + 1
                MSFlexGrid.Col = i2
                If fld.name = "Comissao" Then
                    If fld.Value = True Then
                        MSFlexGrid.Text = "Sim"
                    Else
                        MSFlexGrid.Text = "Não"
                    End If
                Else
                    If IsNull(fld.Value) Then
                        MSFlexGrid.Text = ""
                    Else
                        MSFlexGrid.Text = fld.Value
                    End If
                End If
            Next
            lRst.MoveNext
        Loop
        MSFlexGrid.Visible = True
        MSFlexGrid.Row = 1
        MSFlexGrid.Col = 0
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If txt_condicao.Text = "" Then
        MsgBox "Informe a condição testada.", vbInformation, "Atenção!"
        txt_condicao.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_ok_Click()
    MarcaCelulas
    cmd_sair_Click
    
    'Call GravaConfiguracaoConsulta(Me.name, cbo_campo.Text, cbo_operador.Text, txt_condicao.Text)
    'AtualizaGrid
    'If cbo_campo.Text = "[Codigo de Barra]" Then
    '    txt_condicao.SetFocus
    'Else
    '    MSFlexGrid.SetFocus
    'End If
    'MSFlexGrid.SetFocus
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    If lFlagConsulta = 0 Then
        FormaGrid
        Inicializa
        txt_condicao.SetFocus
    Else
        lFlagConsulta = 0
    End If
End Sub
Private Sub Inicializa()
    Dim i As Integer
    
    txt_condicao.Text = ""
    'If BuscaConfiguracaoConsulta(Me.name) Then
    '    'x_campo = RetiraGString(1)
    '    'x_operador = RetiraGString(2)
    '    txt_condicao.Text = RetiraGString(3)
    '    g_string = ""
    'End If
    'AtualizaGrid
End Sub
Private Sub Form_Deactivate()
    lFlagConsulta = 1
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
    Screen.MousePointer = 1
    CentraForm Me
    'lSQL = "SELECT * FROM Produto WHERE Codigo = 0"
    'Set lRst = New adodb.Recordset
    'Set lRst = Conectar.RsConexao(lSQL)
    'lRst.Close
    g_string = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub MSFlexGrid_DblClick()
    MarcaCelulas
    cmd_sair_Click
End Sub
Private Sub MSFlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        MarcaCelulas
        cmd_sair_Click
    ElseIf KeyCode = 98 Then
        KeyCode = 40
    ElseIf KeyCode = 99 Then
        KeyCode = 34
    ElseIf KeyCode = 104 Then
        KeyCode = 38
    ElseIf KeyCode = 105 Then
        KeyCode = 33
    End If
End Sub
Private Sub MSFlexGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0
        MarcaCelulas
        cmd_sair_Click
    End If
End Sub
Private Sub txt_condicao_GotFocus()
    txt_condicao.SelStart = 0
    txt_condicao.SelLength = Len(txt_condicao.Text)
End Sub
Private Sub txt_condicao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        MSFlexGrid.SetFocus
    End If
End Sub
Private Sub txt_condicao_LostFocus()
    AtualizaGrid
    If MSFlexGrid.Rows > 1 Then
        cmd_ok.SetFocus
    Else
        txt_condicao.SetFocus
    End If
End Sub
