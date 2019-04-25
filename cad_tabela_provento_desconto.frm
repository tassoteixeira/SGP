VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form cadastro_tabela_provento_desconto 
   Caption         =   "Tabela de Proventos/Descontos"
   ClientHeight    =   6255
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   7875
   Icon            =   "cad_tabela_provento_desconto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cad_tabela_provento_desconto.frx":030A
   ScaleHeight     =   6255
   ScaleWidth      =   7875
   Begin VB.Frame frmDados 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7635
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   2040
         MaxLength       =   40
         TabIndex        =   4
         Top             =   600
         Width           =   5475
      End
      Begin VB.TextBox txt_valor 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txt_percentual 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txt_codigo 
         Height          =   285
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   2
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox txt_fracao 
         Height          =   285
         Left            =   5340
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1320
         Width           =   2175
      End
      Begin Threed.SSCheck chk_automatico 
         Height          =   195
         Left            =   2040
         TabIndex        =   15
         Top             =   2040
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "&Automatico"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption opt_provento 
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   1680
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "&Provento"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption opt_desconto 
         Height          =   255
         Left            =   4140
         TabIndex        =   13
         Top             =   1680
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "&Desconto"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Movimento"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   1980
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Nome"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Valor"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Percen&tual"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Códi&go"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Provento/Desconto"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Fração"
         Height          =   315
         Index           =   5
         Left            =   4140
         TabIndex        =   9
         Top             =   1320
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmd_mais 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   2580
      Width           =   435
   End
   Begin VB.Frame frmDados2 
      Height          =   2415
      Left            =   120
      TabIndex        =   18
      Top             =   2820
      Width           =   7635
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
         Height          =   2355
         Left            =   0
         TabIndex        =   32
         Top             =   60
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   4154
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483633
      End
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cad_tabela_provento_desconto.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Cria um novo registro."
      Top             =   5340
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cad_tabela_provento_desconto.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Altera o registro atual."
      Top             =   5340
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cad_tabela_provento_desconto.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Exclui o registro atual."
      Top             =   5340
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "cad_tabela_provento_desconto.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   5340
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "cad_tabela_provento_desconto.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5340
      Width           =   795
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   5580
      TabIndex        =   26
      Top             =   5220
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cad_tabela_provento_desconto.frx":7472
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cad_tabela_provento_desconto.frx":89F4
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cad_tabela_provento_desconto.frx":9E66
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cad_tabela_provento_desconto.frx":B360
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6960
      Picture         =   "cad_tabela_provento_desconto.frx":C85A
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Cancela o registro atual."
      Top             =   5340
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   6060
      Picture         =   "cad_tabela_provento_desconto.frx":DD54
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Confirma o registro atual."
      Top             =   5340
      Width           =   795
   End
   Begin MSAdodcLib.Adodc adodcProventoDesconto 
      Height          =   330
      Left            =   4620
      Top             =   2340
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
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
      Caption         =   "adodcProventoDesconto"
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
   Begin MSDataListLib.DataCombo dtcboProventoDesconto 
      Bindings        =   "cad_tabela_provento_desconto.frx":F35E
      Height          =   315
      Left            =   3600
      TabIndex        =   31
      Top             =   2520
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "Nome"
      BoundColumn     =   "Codigo"
      Text            =   "dtcboProventoDesconto"
   End
   Begin VB.Label Label1 
      Caption         =   "&Base de Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   16
      Top             =   2580
      Width           =   2475
   End
End
Attribute VB_Name = "cadastro_tabela_provento_desconto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lCodigo As Integer
Private TabelaProventoDesconto As New cTabelaProventoDesconto
Private Sub AdcionaDadosGridBaseCalculo(x_codigo As Integer, x_nome As String)
    Dim i As Integer
    MSFlexGrid.Rows = MSFlexGrid.Rows + 1
    i = MSFlexGrid.Rows - 2
    MSFlexGrid.Row = i
    MSFlexGrid.Col = 0
    MSFlexGrid.Text = Format(x_codigo, "#000")
    MSFlexGrid.Col = 1
    MSFlexGrid.Text = x_nome
End Sub
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    txt_codigo.Enabled = True
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_codigo = 1
    If TabelaProventoDesconto.LocalizarUltimo Then
        txt_codigo.Text = TabelaProventoDesconto.Codigo + 1
    End If
    'With tbl_tabela_provento_desconto
    '    If .RecordCount > 0 Then
    '        .Seek "<", 9999
    '        If Not .NoMatch Then
    '            txt_codigo = !Codigo + 1
    '        End If
    '    End If
    'End With
End Sub
Private Sub AtualTabe()
    Dim i As Integer
    TabelaProventoDesconto.Codigo = Val(txt_codigo.Text)
    TabelaProventoDesconto.Nome = txt_nome.Text
    TabelaProventoDesconto.valor = fValidaValor2(txt_valor.Text)
    TabelaProventoDesconto.Fracao = txt_fracao.Text
    TabelaProventoDesconto.Percentual = fValidaValor2(txt_percentual.Text)
    If opt_provento.Value = True Then
        TabelaProventoDesconto.ProventoouDesconto = "P"
    Else
        TabelaProventoDesconto.ProventoouDesconto = "D"
    End If
    TabelaProventoDesconto.Automatico = chk_automatico.Value
    TabelaProventoDesconto.BaseparaCalculo = ""

    For i = 1 To (MSFlexGrid.Rows - 2)
        If MSFlexGrid.TextMatrix(i, 0) <> "" Then
            If Val(MSFlexGrid.TextMatrix(i, 0)) > 0 Then
                TabelaProventoDesconto.BaseparaCalculo = TabelaProventoDesconto.BaseparaCalculo + Format(Val(MSFlexGrid.TextMatrix(i, 0)), "0000") + "@"
            End If
        End If
    Next
End Sub
Private Sub AtualTela()
    Dim i As Integer
    Dim i2 As Integer
    Dim x_codigo As Integer
    Dim x_nome As String
    Dim xString As String
    lCodigo = TabelaProventoDesconto.Codigo
    txt_codigo.Text = Format(TabelaProventoDesconto.Codigo, "#000")
    txt_nome.Text = TabelaProventoDesconto.Nome
    txt_valor.Text = Format(TabelaProventoDesconto.valor, "###,##0.00")
    txt_percentual.Text = Format(TabelaProventoDesconto.Percentual, "##0.00")
    txt_fracao.Text = TabelaProventoDesconto.Fracao
    If TabelaProventoDesconto.ProventoouDesconto = "P" Then
        opt_provento.Value = True
    Else
        opt_desconto.Value = True
    End If
    chk_automatico.Value = TabelaProventoDesconto.Automatico
    
    'Monta Grid de Base de Calculo
    LimpaMSFlexGrid
    xString = TabelaProventoDesconto.BaseparaCalculo
    If Len(xString) > 0 Then
        i2 = Len(xString) / 5
        i = 1
        Do Until i > i2
            x_codigo = RetiraString(xString, i)
            If TabelaProventoDesconto.LocalizarCodigo(x_codigo) Then
                x_nome = TabelaProventoDesconto.Nome
            Else
                x_nome = "** Não Cadastrado **"
            End If
            If Not ExisteItemGridBaseCalculo(x_codigo) Then
                Call AdcionaDadosGridBaseCalculo(x_codigo, x_nome)
            Else
                MsgBox "Já existe este provento.", vbInformation, "Atenção!"
            End If
            i = i + 1
        Loop
        Call TabelaProventoDesconto.LocalizarCodigo(Val(txt_codigo.Text))
    End If
    frmDados.Enabled = False
    frmDados2.Enabled = False
    cmd_mais.Enabled = False
End Sub
'Function BuscaRegistro(x_codigo As Integer) As Boolean
'    BuscaRegistro = False
'    If tbl_tabela_provento_desconto.RecordCount > 0 Then
'        tbl_tabela_provento_desconto.Seek "=", x_codigo
'        If Not tbl_tabela_provento_desconto.NoMatch Then
'            AtualTela
'            BuscaRegistro = True
'            Exit Function
'        End If
'    End If
'End Function
'Function BuscaDados() As Boolean
'    BuscaDados = False
'    With tbl_tabela_provento_desconto
'        If .RecordCount > 0 Then
'            If lOpcao = 3 Then
'                If Not .EOF Then
'                    .MoveNext
'                    If Not .EOF Then
'                        AtualTela
'                        BuscaDados = True
'                        Exit Function
'                    End If
'                End If
'            End If
'            .Seek "<", 9999
'            If Not .NoMatch Then
'                AtualTela
'                BuscaDados = True
'                Exit Function
'            End If
'        End If
'        LimpaTela
'    End With
'End Function
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_pesquisa.Enabled = False
    cmd_excluir.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub Finaliza()
    Set TabelaProventoDesconto = Nothing
    frm_cadastro.Show
End Sub
Private Sub PesquisaBaseCalculo()
    'LimpaMSFlexGrid
    'With tbl_tabela_provento_desconto
    '    .Seek ">=", lCodigo
    '    If Not .NoMatch Then
    '        Do Until .EOF
    '            AdcionaDadosGridBaseCalculo
    '            .MoveNext
    '        Loop
    '    End If
    '    Call BuscaRegistro(lCodigo)
    '    grid_base_calculo.Row = grid_base_calculo.Rows - 1
    '    grid_base_calculo.Col = 0
    'End With
End Sub
Private Sub dtcboProventoDesconto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        MSFlexGrid.SetFocus
    End If
End Sub
Private Sub dtcboProventoDesconto_LostFocus()
    If dtcboProventoDesconto.BoundText <> "" Then
        If Not ExisteItemGridBaseCalculo(Val(dtcboProventoDesconto.BoundText)) Then
            Call AdcionaDadosGridBaseCalculo(Val(dtcboProventoDesconto.BoundText), dtcboProventoDesconto.Text)
        Else
            MsgBox "Já existe este provento.", vbInformation, "Atenção!"
        End If
    End If
    dtcboProventoDesconto.Visible = False
End Sub
Private Sub chk_automatico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub cmd_alterar_Click()
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frmDados.Enabled = True
    frmDados2.Enabled = True
    cmd_mais.Enabled = True
    txt_codigo.Enabled = False
    txt_valor.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If TabelaProventoDesconto.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If TabelaProventoDesconto.LocalizarCodigo(lCodigo) Then
        AtualTela
        AtivaBotoes
        cmd_alterar.SetFocus
    Else
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
        cmd_novo.SetFocus
    End If
End Sub
Private Sub LimpaMSFlexGrid()
    Dim i As Integer
    MSFlexGrid.WordWrap = True
    MSFlexGrid.Cols = 2
    MSFlexGrid.Rows = 2
    MSFlexGrid.Row = 1
    For i = 0 To (MSFlexGrid.Cols - 1)
        MSFlexGrid.Col = i
        MSFlexGrid.Text = ""
    Next
    MSFlexGrid.RowHeight(0) = 500
    MSFlexGrid.Row = 0
    i = 0
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Código"
    MSFlexGrid.ColWidth(i) = 700
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Provento/Desconto"
    MSFlexGrid.ColWidth(i) = 3000
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
End Sub
Private Sub LimpaTela()
    txt_codigo.Text = ""
    txt_nome.Text = ""
    txt_valor.Text = ""
    txt_percentual.Text = ""
    txt_fracao.Text = ""
    opt_provento.Value = False
    opt_desconto.Value = True
    chk_automatico.Value = False
    LimpaMSFlexGrid
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_codigo.Text) > 0 Then
        If (MsgBox("Deseja realmente excluir este registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If TabelaProventoDesconto.Excluir(Val(txt_codigo.Text)) Then
                LimpaTela
                If TabelaProventoDesconto.LocalizarUltimo Then
                    AtualTela
                Else
                    DesativaBotoes
                    cmd_novo.Enabled = True
                    cmd_sair.Enabled = True
                    cmd_novo.SetFocus
                End If
            Else
                MsgBox "Não foi possivel excluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
    End If
End Sub
Private Sub cmd_mais_Click()
    dtcboProventoDesconto.Visible = True
    adodcProventoDesconto.ConnectionString = gConnectionString
    adodcProventoDesconto.RecordSource = "Select * From Tabela_Provento_Desconto Where Codigo <> " & Val(txt_codigo) & " Order By Nome"
    adodcProventoDesconto.Refresh
    dtcboProventoDesconto.SetFocus
End Sub
Private Sub cmd_novo_Click()
    LimpaTela
    Inclui
    frmDados.Enabled = True
    frmDados2.Enabled = True
    cmd_mais.Enabled = True
    txt_nome.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If TabelaProventoDesconto.Incluir Then
                lCodigo = Val(txt_codigo.Text)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not TabelaProventoDesconto.Alterar(lCodigo) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        lOpcao = 0
        Call TabelaProventoDesconto.LocalizarCodigo(lCodigo)
        AtualTela
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_tabela_provento_desconto.Name, "Provento/Descontoo"
    Exit Sub
End Sub
Function RetiraString(x_string As String, numero As Integer) As Integer
    RetiraString = 0
    Dim x_index As Integer
    Dim x_inicio As Integer
    Dim x_numero As Integer
    x_inicio = 1
    x_numero = 1
    If Len(x_string) > 0 Then
        Do Until x_index > Len(x_string)
            x_index = x_index + 1
            If Mid(x_string, x_index, 1) = "@" Then
                If x_numero = numero Then
                    RetiraString = Mid(x_string, x_inicio, x_index - x_inicio)
                    Exit Function
                End If
                'x_index = x_index + 2
                x_numero = x_numero + 1
                x_inicio = x_index + 1
            End If
        Loop
    End If
End Function
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not Val(txt_codigo) > 0 Then
        MsgBox "Informe o código.", vbInformation, "Atenção!"
        txt_codigo.SetFocus
    ElseIf Not txt_nome <> "" Then
        MsgBox "Informe o nome do provento/desconto.", vbInformation, "Atenção!"
        txt_nome.SetFocus
    ElseIf fValidaValor2(txt_valor) = 0 And fValidaValor2(txt_percentual) = 0 And txt_fracao = "" Then
        MsgBox "Informe o valor, percentual ou fração.", vbInformation, "Atenção!"
        txt_valor.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_tabela_provento_desconto.Show 1
    If Len(g_string) > 0 Then
        lCodigo = RetiraGString(1)
        If TabelaProventoDesconto.LocalizarCodigo(lCodigo) Then
            AtualTela
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If TabelaProventoDesconto.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If TabelaProventoDesconto.LocalizarProximo Then
        AtualTela
    Else
        MsgBox "Fim de Arquivo.", vbInformation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If TabelaProventoDesconto.LocalizarUltimo Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        DesativaBotoes
        If TabelaProventoDesconto.LocalizarUltimo Then
            AtualTela
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        cmd_novo.SetFocus
    Else
        lFlagCadastro = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_Deactivate()
    lFlagCadastro = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And lOpcao = 0 Then
        KeyCode = 0
        cmd_novo_Click
    ElseIf KeyCode = vbKeyF3 And lOpcao = 0 Then
        KeyCode = 0
        cmd_alterar_Click
    ElseIf KeyCode = vbKeyF4 And Shift = 0 Then
        KeyCode = 0
        cmd_excluir_Click
    ElseIf KeyCode = vbKeyF5 And lOpcao = 0 Then
        KeyCode = 0
        cmd_pesquisa_Click
    ElseIf KeyCode = vbKeyF7 And lOpcao = 0 Then
        KeyCode = 0
        cmd_primeiro_Click
    ElseIf KeyCode = vbKeyF8 And lOpcao = 0 Then
        KeyCode = 0
        cmd_anterior_Click
    ElseIf KeyCode = vbKeyF9 And lOpcao = 0 Then
        KeyCode = 0
        cmd_proximo_Click
    ElseIf KeyCode = vbKeyF10 And lOpcao = 0 Then
        KeyCode = 0
        cmd_ultimo_Click
    ElseIf KeyCode = vbKeyF11 And lOpcao > 0 Then
        KeyCode = 0
        cmd_ok_Click
    ElseIf KeyCode = vbKeyF12 And lOpcao > 0 Then
        KeyCode = 0
        cmd_cancelar_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub MarcaCelulaBaseCalculo()
    If MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0) <> "" Then
        If (MsgBox("Deseja excluir este provento/desconto da base de cálculo?", 4 + 32 + 256, "Exclusão de Base de Cálculo!")) = 6 Then
            MSFlexGrid.Col = 0
            MSFlexGrid.Text = ""
            MSFlexGrid.Col = 1
            MSFlexGrid.Text = ""
        End If
    End If
End Sub
Function ExisteItemGridBaseCalculo(x_codigo As Integer) As Boolean
    ExisteItemGridBaseCalculo = False
    Dim i As Integer
    For i = 1 To (MSFlexGrid.Rows - 2)
        If MSFlexGrid.TextMatrix(i, 0) <> "" Then
            If Val(MSFlexGrid.TextMatrix(i, 0)) = x_codigo Then
                ExisteItemGridBaseCalculo = True
                Exit Function
            End If
        End If
    Next
End Function
Private Sub MSFlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        MarcaCelulaBaseCalculo
    End If
End Sub
Private Sub opt_desconto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_automatico.SetFocus
    End If
End Sub
Private Sub opt_provento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_automatico.SetFocus
    End If
End Sub
Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_nome.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_codigo_LostFocus()
    txt_codigo.Text = Format(txt_codigo.Text, "#000")
    If lOpcao = 1 And Val(txt_codigo.Text) > 0 Then
        If TabelaProventoDesconto.LocalizarCodigo(Val(txt_codigo.Text)) Then
            MsgBox "Já existe Provento/Desconto cadastrado com este código." & Chr(10) & Chr(10) & "Mude o código informado.", vbInformation, "Duplicidade de Registro!"
            txt_codigo.SetFocus
        End If
    End If
End Sub
Private Sub txt_fracao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If opt_provento.Value = True Then
            opt_provento.SetFocus
        Else
            opt_desconto.SetFocus
        End If
    End If
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor.SetFocus
    End If
End Sub
Private Sub txt_nome_LostFocus()
    If lOpcao = 1 And txt_nome.Text <> "" Then
        If TabelaProventoDesconto.LocalizarNome(txt_nome.Text) Then
            If (MsgBox("Já existe Provento/Desconto cadastrado com este nome." & Chr(10) & Chr(10) & "Deseja cadastrar assim mesmo?", 4 + 32 + 256, "Duplicidade de Registro!")) = 7 Then
                txt_nome.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txt_percentual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_fracao.SetFocus
    End If
End Sub
Private Sub txt_percentual_LostFocus()
    txt_percentual = Format(txt_percentual, "##0.00")
End Sub
Private Sub txt_valor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_percentual.SetFocus
    End If
End Sub
Private Sub txt_valor_LostFocus()
    txt_valor = Format(txt_valor, "###,##0.00")
End Sub
