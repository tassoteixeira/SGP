VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form ConsultaEstoque 
   Caption         =   "Consulta de Estoque"
   ClientHeight    =   8295
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   8550
   Icon            =   "ConsultaEstoque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "ConsultaEstoque.frx":0442
   ScaleHeight     =   8295
   ScaleWidth      =   8550
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   7620
      Picture         =   "ConsultaEstoque.frx":0888
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   7380
      Width           =   795
   End
   Begin VB.Frame frmDados 
      Enabled         =   0   'False
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8295
      Begin VB.TextBox txt_produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   120
         MaxLength       =   18
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc adodcProduto 
         Height          =   330
         Left            =   2280
         Top             =   360
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adodcProduto"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo dtcboProduto 
         Bindings        =   "ConsultaEstoque.frx":1F1A
         Height          =   420
         Left            =   1260
         TabIndex        =   3
         Top             =   360
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   741
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ForeColor       =   16711680
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboProduto"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txt_preview 
         Height          =   6375
         Left            =   60
         TabIndex        =   6
         Top             =   780
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   11245
         _Version        =   393217
         BackColor       =   16777215
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         RightMargin     =   9,99999e5
         TextRTF         =   $"ConsultaEstoque.frx":1F35
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "P&roduto"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   150
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   6720
      Picture         =   "ConsultaEstoque.frx":1FB7
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Confirma o registro atual."
      Top             =   7380
      Width           =   795
   End
End
Attribute VB_Name = "ConsultaEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lEmpresa As Integer
Dim lSQL As String

Dim rsDados As New adodb.Recordset

Private Aliquota As New cAliquota
Private EntradaProduto As New cEntradaProduto
Private Estoque As New cEstoque
Private Fornecedor As New cFornecedor
Private Grupo As New cGrupo
Private MovimentoLubrificante As New cMovimentoLubrificante
Private Produto As New cProduto
Private SubEstoque As New cSubEstoque
Private TipoSubEstoque As New cTipoSubEstoque

Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Aliquota = Nothing
    Set EntradaProduto = Nothing
    Set Estoque = Nothing
    Set Fornecedor = Nothing
    Set Grupo = Nothing
    Set MovimentoLubrificante = Nothing
    Set Produto = Nothing
    Set SubEstoque = Nothing
    Set TipoSubEstoque = Nothing
End Sub
Private Sub LimpaTela()
    txt_produto.Text = ""
    dtcboProduto.BoundText = ""
    Me.txt_preview.Text = ""
End Sub
Private Sub cmd_ok_Click()
    If dtcboProduto.BoundText <> "" Then
        Call GravaAuditoria(1, Me.name, 10, "Produto:" & txt_produto.Text & " - " & dtcboProduto.Text)
        MontaDados
    End If
End Sub
Private Sub MontaDados()
    txt_preview.Text = ""
    MontaDadosProduto
End Sub
Private Sub MontaDadosProduto()
    Dim xString As String
    Dim i As Integer
    Dim xValor As Currency
    Dim xDataInicial As Date
    Dim xDataFinal As Date
                
    xDataInicial = Date
    xDataFinal = Date
    xString = Space(100)
    Mid(xString, 1, 20) = "UNIDADE............:"
    Mid(xString, 22, 20) = Produto.Unidade
    txt_preview.Text = txt_preview.Text & xString & Chr(10)
    
    xString = Space(100)
    Mid(xString, 1, 20) = "PREÇO DE CUSTO.....:"
    Mid(xString, 22, 20) = Format(Produto.PrecoCusto, "###,##0.00")
    txt_preview.Text = txt_preview.Text & xString & Chr(10)
    
    If Estoque.LocalizarCodigo(g_empresa, Produto.Codigo) Then
        xString = Space(100)
        Mid(xString, 1, 20) = "PREÇO DE VENDA.....:"
        Mid(xString, 22, 20) = Format(Estoque.PrecoVenda, "###,##0.00")
        txt_preview.Text = txt_preview.Text & xString & Chr(10)
    
        xString = Space(100)
        Mid(xString, 1, 20) = "MARGEM DE LUCRO....:"
        xValor = (Estoque.PrecoVenda - Produto.PrecoCusto) * 100 / Produto.PrecoCusto
        Mid(xString, 22, 20) = Format(xValor, "###,##0.0000") & "%"
        txt_preview.Text = txt_preview.Text & xString & Chr(10)
    End If

    If Grupo.LocalizarCodigo(Produto.CodigoGrupo) Then
        xString = Space(100)
        Mid(xString, 1, 20) = "GRUPO..............:"
        Mid(xString, 22, 40) = Grupo.Nome
        txt_preview.Text = txt_preview.Text & xString & Chr(10)
    End If
    
    If Aliquota.LocalizarCodigoAliquota(Produto.CodigoAliquota) Then
        xString = Space(100)
        Mid(xString, 1, 20) = "ALÍQUOTA...........:"
        Mid(xString, 22, 40) = Aliquota.Nome
        txt_preview.Text = txt_preview.Text & xString & Chr(10)
    End If
    
    txt_preview.Text = txt_preview.Text & Chr(10)
    xString = Space(100)
    Mid(xString, 22, 20) = "ESTOQUES"
    txt_preview.Text = txt_preview.Text & xString & Chr(10)
    For i = 1 To 10
        If TipoSubEstoque.LocalizarCodigo(i) Then
            xString = Space(100)
            Mid(xString, 1, 20) = "...................:"
            Mid(xString, 1, 19) = Mid(TipoSubEstoque.Nome, 1, 19)
            If SubEstoque.LocalizarCodigo(g_empresa, Produto.Codigo, i) Then
                Mid(xString, 22, 20) = Format(SubEstoque.Quantidade, "###,##0.00")
            End If
            txt_preview.Text = txt_preview.Text & xString & Chr(10)
        Else
            Exit For
        End If
    Next
    If Estoque.LocalizarCodigo(g_empresa, Produto.Codigo) Then
        xString = Space(100)
        Mid(xString, 1, 20) = "TOTAL GERAL........:"
        Mid(xString, 22, 20) = Format(Estoque.Quantidade, "###,##0.00")
        txt_preview.Text = txt_preview.Text & xString & Chr(10)
    End If
    
    txt_preview.Text = txt_preview.Text & Chr(10)
    xString = Space(100)
    Mid(xString, 22, 20) = "ÚLTIMA ENTRADA"
    txt_preview.Text = txt_preview.Text & xString & Chr(10)
    xString = Space(100)
    If EntradaProduto.LocalizarUltimoProduto(g_empresa, Produto.Codigo) Then
        Mid(xString, 1, 20) = "FORNECEDOR.........:"
        If Fornecedor.LocalizarCodigo(g_empresa, EntradaProduto.CodigoFornecedor) Then
            Mid(xString, 22, 40) = Fornecedor.Nome
        End If
        txt_preview.Text = txt_preview.Text & xString & Chr(10)
    
        xString = Space(100)
        Mid(xString, 1, 20) = "DATA...............:"
        Mid(xString, 22, 20) = Format(EntradaProduto.DataEntrada, "dd/mm/yyyy")
        xDataInicial = EntradaProduto.DataEntrada
        txt_preview.Text = txt_preview.Text & xString & Chr(10)
    
        xString = Space(100)
        Mid(xString, 1, 20) = "NUMERO DA NOTA.....:"
        Mid(xString, 22, 20) = EntradaProduto.NumeroDocumento
        txt_preview.Text = txt_preview.Text & xString & Chr(10)
    
        xString = Space(100)
        Mid(xString, 1, 20) = "VALOR UNITÁRIO.....:"
        Mid(xString, 22, 20) = Format(EntradaProduto.PrecoCusto, "###,##0.00")
        txt_preview.Text = txt_preview.Text & xString & Chr(10)
    
        xString = Space(100)
        Mid(xString, 1, 20) = "QUANTIDADE.........:"
        Mid(xString, 22, 20) = Format(EntradaProduto.Quantidade, "###,##0.00")
        txt_preview.Text = txt_preview.Text & xString & Chr(10)
    
        xString = Space(100)
        Mid(xString, 1, 20) = "TOTAL DO CUSTO.....:"
        Mid(xString, 22, 20) = Format(EntradaProduto.TotalCusto, "###,##0.00")
        txt_preview.Text = txt_preview.Text & xString & Chr(10)
    Else
        Mid(xString, 1, 20) = "NÃO TEM ENTRADA"
        txt_preview.Text = txt_preview.Text & xString & Chr(10)
    End If
    
    xValor = MovimentoLubrificante.TotalQtd(g_empresa, xDataInicial, xDataFinal, Produto.Codigo)
    xString = Space(100)
    Mid(xString, 1, 30) = "VENDAS DESDE A ÚLTIMA ENTRADA:"
    Mid(xString, 32, 20) = Format(xValor, "###,##0.00")
    txt_preview.Text = txt_preview.Text & xString & Chr(10)
    If xValor > 0 Then
        i = DateDiff("d", xDataInicial, xDataFinal) + 1
        xValor = xValor / i
        xString = Space(100)
        Mid(xString, 1, 30) = "MÉDIA DE VENDA DIÁRIA........:"
        Mid(xString, 32, 20) = Format(xValor, "###,##0.00")
        txt_preview.Text = txt_preview.Text & xString & Chr(10)
    End If
    xValor = MovimentoLubrificante.TotalQtd(g_empresa, Date, Date, Produto.Codigo)
    xString = Space(100)
    Mid(xString, 1, 30) = "QUANTIDADE VENDIDA HOJE......:"
    Mid(xString, 32, 20) = Format(xValor, "###,##0.00")
    txt_preview.Text = txt_preview.Text & xString & Chr(10)

    txt_preview.SelStart = 0
    txt_preview.SelLength = Len(txt_preview.Text)
    txt_preview.SelColor = 16711680
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub dtcboProduto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmd_ok.SetFocus
    End If
End Sub
Private Sub dtcboProduto_LostFocus()
    If dtcboProduto.BoundText <> "" Then
        txt_produto.Text = dtcboProduto.BoundText
        txt_produto_LostFocus
        cmd_ok_Click
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    frmDados.Enabled = True
    txt_produto.SetFocus
    Screen.MousePointer = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        cmd_sair_Click
    ElseIf KeyCode = vbKeyF11 Then
        KeyCode = 0
        cmd_ok_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    Set adodcProduto.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Produto WHERE Inativo = " & preparaBooleano(False) & " ORDER BY Nome")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub txt_produto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboProduto.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_produto_LostFocus()
    If Val(txt_produto.Text) > 0 Then
        If Len(txt_produto.Text) > 5 Then
            If Produto.LocalizarCodigoBarra(txt_produto.Text) Then
                txt_produto.Text = Produto.Codigo
            Else
                MsgBox "Codigo de Barra não cadastrado!", vbInformation, "Erro de Leitura!"
                txt_produto.Text = ""
                txt_produto.SetFocus
                Exit Sub
            End If
        End If
        If Produto.LocalizarCodigo(CLng(txt_produto.Text)) Then
            dtcboProduto.BoundText = CLng(txt_produto.Text)
            cmd_ok.SetFocus
        Else
            MsgBox "Produto não cadastrado.", vbInformation, "Atenção!"
            txt_produto.SetFocus
            Exit Sub
        End If
    End If
End Sub

