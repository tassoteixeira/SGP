VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form movimento_previsao_venda_prazo 
   Caption         =   "Movimento de Previsão de Venda à Prazo"
   ClientHeight    =   7260
   ClientLeft      =   1410
   ClientTop       =   1545
   ClientWidth     =   8955
   Icon            =   "movimento_previsao_venda_prazo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "movimento_previsao_venda_prazo.frx":030A
   ScaleHeight     =   7260
   ScaleWidth      =   8955
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "movimento_previsao_venda_prazo.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cria um novo registro."
      Top             =   6300
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "movimento_previsao_venda_prazo.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Altera o registro atual."
      Top             =   6300
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "movimento_previsao_venda_prazo.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Exclui o registro atual."
      Top             =   6300
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "movimento_previsao_venda_prazo.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   6300
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "movimento_previsao_venda_prazo.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   6300
      Width           =   795
   End
   Begin VB.Frame frmDados 
      Enabled         =   0   'False
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8715
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   3180
         Picture         =   "movimento_previsao_venda_prazo.frx":7472
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_celula 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   4800
         TabIndex        =   18
         Top             =   2940
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   300
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid fgd_composicao_caixa 
         Height          =   4275
         Left            =   0
         TabIndex        =   19
         Top             =   1320
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   7541
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorBkg    =   12632256
         AllowUserResizing=   1
      End
      Begin MSAdodcLib.Adodc adodc_combustivel 
         Height          =   330
         Left            =   2700
         Top             =   720
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "adodc_combustivel"
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
      Begin MSDataListLib.DataCombo dtcbo_combustivel 
         Bindings        =   "movimento_previsao_venda_prazo.frx":874C
         Height          =   315
         Left            =   2040
         TabIndex        =   5
         Top             =   720
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_combustivel"
      End
      Begin VB.Label Label6 
         Caption         =   "Total da Venda à Prazo do Dia"
         Height          =   315
         Left            =   3480
         TabIndex        =   21
         Top             =   5700
         Width           =   2415
      End
      Begin VB.Label lbl_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5940
         TabIndex        =   20
         Top             =   5700
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Co&mbustível"
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Data do movimento"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   6660
      TabIndex        =   13
      Top             =   6180
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "movimento_previsao_venda_prazo.frx":876C
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "movimento_previsao_venda_prazo.frx":9C66
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "movimento_previsao_venda_prazo.frx":B160
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "movimento_previsao_venda_prazo.frx":C5D2
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   7140
      Picture         =   "movimento_previsao_venda_prazo.frx":DB54
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Confirma o registro atual."
      Top             =   6300
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   8040
      Picture         =   "movimento_previsao_venda_prazo.frx":F15E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cancela o registro atual."
      Top             =   6300
      Width           =   795
   End
End
Attribute VB_Name = "movimento_previsao_venda_prazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagMovimento As Integer
Dim lOpcao As String
Dim lEmpresa As Integer
Dim lData As Date
Dim lTipoCombustivel As String

Const NovaLinha As String = ">*"      ' Indica uma nova linha
Private ControlVisible As Boolean     ' Se o controle esta visivel ou nao
Private LastRow As Long               ' Ultima linha em que se editou
Private LastCol As Long               ' ultima coluna em que se editou
Dim lMarcaCelula As Boolean

Private rsPrevisaoVendaPrazo As New adodb.Recordset
Private PrevisaoVendaPrazo As New cPrevisaoVendaPrazo
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    If g_nivel_acesso > 4 Then
        If g_empresa < g_cfg_empresa_i Or g_empresa > g_cfg_empresa_f Then
            cmd_novo.Enabled = False
            cmd_alterar.Enabled = False
            cmd_excluir.Enabled = False
        End If
    End If
    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Private Sub AtribuiValorCelula()
    Dim Texto As String
    '
    txt_celula.Visible = False
    ControlVisible = False
    '
    ' atribuir o texto anterior a celula
    Select Case LastCol
      Case 4 To 7
        'notas menores que 5 muda cor fonte para vermelho, demais azul
        Texto = txt_celula.Text
        fgd_composicao_caixa.TextMatrix(LastRow, LastCol) = Texto
        'If Val(fgd_composicao_caixa.Text) < 6 Then
        '     fgd_composicao_caixa.CellForeColor = vbRed
        'Else
        '     fgd_composicao_caixa.CellForeColor = vbBlue
        'End If
      Case Else
        'If LastRow = 0 And LastCol = 0 Then
            LastRow = fgd_composicao_caixa.Row
            LastCol = fgd_composicao_caixa.Col
        'End If
      
        Texto = txt_celula.Text
        fgd_composicao_caixa.TextMatrix(LastRow, LastCol) = Texto
    End Select
End Sub
Private Sub AtualizaGrid()
    Dim i As Integer
    Dim xSQL As String
    
    LimpaGrid
    i = 0
    fgd_composicao_caixa.Visible = False
    
    xSQL = ""
    xSQL = xSQL & "   SELECT Data, [Tipo de Combustivel], [Previsao de Venda a Prazo], [Media de Venda Diaria a Prazo], "
    xSQL = xSQL & "          [Total da Venda], [Quantidade de Venda a Prazo], Saldo, Hora"
    xSQL = xSQL & "     FROM Previsao_Venda_Prazo"
    xSQL = xSQL & "    WHERE Empresa = " & g_empresa
    xSQL = xSQL & "      AND [Tipo de Combustivel] = " & Chr(39) & dtcbo_combustivel.BoundText & Chr(39)
    xSQL = xSQL & "      AND Data = #" & Format(msk_data.Text, "mm/dd/yyyy") & "#"
    xSQL = xSQL & " ORDER BY Hora"
    Set rsPrevisaoVendaPrazo = New adodb.Recordset
    Set rsPrevisaoVendaPrazo = Conectar.RsConexao(xSQL)
    If Not rsPrevisaoVendaPrazo.EOF Then
        Do Until rsPrevisaoVendaPrazo.EOF
            i = i + 1
            fgd_composicao_caixa.Rows = fgd_composicao_caixa.Rows + 1
            fgd_composicao_caixa.Row = i
            fgd_composicao_caixa.Col = 0
            fgd_composicao_caixa.Text = rsPrevisaoVendaPrazo("Hora").Value
            fgd_composicao_caixa.Col = 1
            fgd_composicao_caixa.Text = Format(rsPrevisaoVendaPrazo("Previsao de Venda a Prazo").Value, "###,###,##0.00")
            fgd_composicao_caixa.Col = 2
            fgd_composicao_caixa.Text = Format(rsPrevisaoVendaPrazo("Media de Venda Diaria a Prazo").Value, "###,###,##0.00")
            fgd_composicao_caixa.Col = 3
            fgd_composicao_caixa.Text = Format(rsPrevisaoVendaPrazo("Total da Venda").Value, "###,###,##0.00")
            fgd_composicao_caixa.Col = 4
            fgd_composicao_caixa.Text = Format(rsPrevisaoVendaPrazo("Quantidade de Venda a Prazo").Value, "###,###,##0.00")
            fgd_composicao_caixa.Col = 5
            fgd_composicao_caixa.Text = Format(rsPrevisaoVendaPrazo("Saldo").Value, "###,###,##0.00")
            rsPrevisaoVendaPrazo.MoveNext
        Loop
    End If
    lbl_total.Caption = Format(PrevisaoVendaPrazo.TotalVendaPrazoDia(g_empresa, dtcbo_combustivel.BoundText, CDate(msk_data.Text)), "###,###,##0.00")
    
    rsPrevisaoVendaPrazo.Close
    Set rsPrevisaoVendaPrazo = Nothing
    fgd_composicao_caixa.Row = 1
    fgd_composicao_caixa.Col = 2
    fgd_composicao_caixa.Visible = True
    frmDados.Enabled = False
End Sub
Private Sub AtualTabe()
    Dim i As Integer
    For i = 1 To (fgd_composicao_caixa.Rows - 2)
        If IsDate(fgd_composicao_caixa.TextMatrix(i, 0)) Then
            PrevisaoVendaPrazo.Empresa = g_empresa
            PrevisaoVendaPrazo.TipoCombustivel = dtcbo_combustivel.BoundText
            PrevisaoVendaPrazo.Data = Format(msk_data.Text, "dd/mm/yyyy")
            PrevisaoVendaPrazo.Hora = fgd_composicao_caixa.TextMatrix(i, 0)
            PrevisaoVendaPrazo.PrevisaoVendaPrazo = fValidaValor(fgd_composicao_caixa.TextMatrix(i, 1))
            PrevisaoVendaPrazo.MediaVendaDiariaPrazo = fValidaValor(fgd_composicao_caixa.TextMatrix(i, 2))
            PrevisaoVendaPrazo.TotalVenda = fValidaValor(fgd_composicao_caixa.TextMatrix(i, 3))
            PrevisaoVendaPrazo.QuantidadeVendaPrazo = fValidaValor(fgd_composicao_caixa.TextMatrix(i, 4))
            PrevisaoVendaPrazo.Saldo = fValidaValor(fgd_composicao_caixa.TextMatrix(i, 5))
            If PrevisaoVendaPrazo.Incluir Then
                lData = msk_data.Text
                lTipoCombustivel = dtcbo_combustivel.BoundText
            Else
                MsgBox "Registro não foi gravado!", vbInformation, "Erro Interno"
            End If
        End If
    Next
End Sub
Private Sub AtualTela()
    Dim i As Integer
    lData = PrevisaoVendaPrazo.Data
    lTipoCombustivel = PrevisaoVendaPrazo.TipoCombustivel
    
    msk_data.Text = Format(PrevisaoVendaPrazo.Data, "dd/mm/yyyy")
    dtcbo_combustivel.BoundText = ""
    dtcbo_combustivel.BoundText = PrevisaoVendaPrazo.TipoCombustivel
    frmDados.Enabled = False
    VerificaLiberacaoDigitacao
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_excluir.Enabled = False
    cmd_pesquisa.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub ExibirCelula()
    Static OK As Boolean
    '
    ' Se for celula fixa , sair
    If fgd_composicao_caixa.Col <= fgd_composicao_caixa.FixedCols - 1 Or fgd_composicao_caixa.Row <= fgd_composicao_caixa.FixedRows - 1 Then
       Exit Sub
    End If
    
    If OK Then Exit Sub
    OK = True
    '
    txt_celula.Visible = False
    '
    LastRow = fgd_composicao_caixa.Row
    LastCol = fgd_composicao_caixa.Col
    If LastCol = 0 Then
        txt_celula.MaxLength = 10
    ElseIf LastCol >= 1 Then
        txt_celula.MaxLength = 10
    End If
    
    '
    ' Nova Celula
    'With fgd_composicao_caixa
    '  If .TextMatrix(LastRow, 0) = NovaLinha Then
    '    .Rows = .Rows + 1
    '    .TextMatrix(LastRow, 0) = LastRow
    '    .TextMatrix(.Rows - 1, 0) = NovaLinha
    '  End If
    'End With
    '
    Select Case LastCol
        Case Else
        txt_celula.Move fgd_composicao_caixa.CellLeft - Screen.TwipsPerPixelX, fgd_composicao_caixa.CellTop + 1300 - Screen.TwipsPerPixelY, fgd_composicao_caixa.CellWidth + Screen.TwipsPerPixelX * 2, fgd_composicao_caixa.CellHeight + Screen.TwipsPerPixelY * 2
        txt_celula.Text = fgd_composicao_caixa.Text
        'If Len(fgd_composicao_caixa.Text) = 0 Then
        '   If LastRow > 1 Then
        '       txt_celula.Text = fgd_composicao_caixa.TextMatrix(LastRow - 1, LastCol)
        '   End If
        'End If
        txt_celula.Visible = True
        If txt_celula.Visible Then
          txt_celula.ZOrder
          txt_celula.SetFocus
        End If
    End Select
    ControlVisible = True
    OK = False
End Sub
Private Sub Finaliza()
    Set PrevisaoVendaPrazo = Nothing
End Sub
Private Sub cmd_alterar_Click()
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frmDados.Enabled = True
    fgd_composicao_caixa.Col = 2
    fgd_composicao_caixa.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If PrevisaoVendaPrazo.LocalizarAnterior Then
        AtualTela
        AtualizaGrid
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If PrevisaoVendaPrazo.LocalizarCodigo(g_empresa, lTipoCombustivel, lData) Then
        AtualTela
        AtualizaGrid
        AtivaBotoes
        cmd_alterar.SetFocus
    Else
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
        cmd_novo.SetFocus
    End If
    lOpcao = 0
End Sub
Private Sub LimpaTela()
    'msk_data.Text = "__/__/____"
    dtcbo_combustivel.BoundText = ""
    LimpaGrid
End Sub
Private Sub LimpaGrid()
    Dim x_sql As String
    Dim i As Integer
    fgd_composicao_caixa.WordWrap = True
    fgd_composicao_caixa.Rows = 2
    fgd_composicao_caixa.Row = 1
    For i = 0 To 5
        fgd_composicao_caixa.Col = i
        fgd_composicao_caixa.Text = ""
    Next
    fgd_composicao_caixa.RowHeight(0) = 750
    fgd_composicao_caixa.Row = 0
    i = 0
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Hora"
    fgd_composicao_caixa.ColWidth(i) = 1110
    fgd_composicao_caixa.ColAlignment(i) = 4
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Previsão de Venda à Prazo"
    fgd_composicao_caixa.ColWidth(i) = 1500
    fgd_composicao_caixa.ColAlignment(i) = 7
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Percentual de Venda Diária à Prazo"
    fgd_composicao_caixa.ColWidth(i) = 1500
    fgd_composicao_caixa.ColAlignment(i) = 7
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Total da Venda"
    fgd_composicao_caixa.ColWidth(i) = 1500
    fgd_composicao_caixa.ColAlignment(i) = 7
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Quantidade de Venda à Prazo"
    fgd_composicao_caixa.ColWidth(i) = 1500
    fgd_composicao_caixa.ColAlignment(i) = 7
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Saldo"
    fgd_composicao_caixa.ColWidth(i) = 1500
    fgd_composicao_caixa.ColAlignment(i) = 7
    txt_celula.Visible = False
    fgd_composicao_caixa.Row = 1
    fgd_composicao_caixa.Col = 0
    fgd_composicao_caixa.Text = ""
    lbl_total.Caption = ""
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data.Text = RetiraGString(1)
    dtcbo_combustivel.SetFocus
    g_string = ""
End Sub
Private Sub cmd_excluir_Click()
    If IsDate(msk_data.Text) Then
        If (MsgBox("Deseja excluir estes registros?", 4 + 32 + 256, "Exclusão de Registros!")) = 6 Then
            If PrevisaoVendaPrazo.ExcluiRegistros(g_empresa, dtcbo_combustivel.BoundText, CDate(msk_data.Text)) Then
                LimpaTela
                If PrevisaoVendaPrazo.LocalizarUltimo(g_empresa) Then
                    AtualTela
                    AtualizaGrid
                Else
                    DesativaBotoes
                    cmd_novo.Enabled = True
                    cmd_sair.Enabled = True
                    cmd_novo.SetFocus
                End If
            Else
                MsgBox "Registros não excluidos!", vbInformation, "Erro de Integridade!"
            End If
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
    LimpaTela
    Inclui
    frmDados.Enabled = True
    msk_data.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        If VerificaLiberacaoDigitacao2 Then
            AtivaBotoes
            If lOpcao = 1 Then
                AtualTabe
            ElseIf lOpcao = 2 Then
                Call PrevisaoVendaPrazo.ExcluiRegistros(g_empresa, lTipoCombustivel, lData)
                AtualTabe
            End If
            lOpcao = 0
            Call PrevisaoVendaPrazo.LocalizarCodigo(g_empresa, lTipoCombustivel, lData)
            AtualizaGrid
            cmd_novo.SetFocus
        End If
    End If
    Exit Sub
FileError:
    MsgBox Error
    Exit Sub
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data do movimento.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf dtcbo_combustivel.BoundText = "" Then
        MsgBox "Escolha um tipo de combustível.", vbInformation, "Atenção!"
        dtcbo_combustivel.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub VerificaLiberacaoDigitacao()
    Dim x_flag As Boolean
    x_flag = True
    If g_nivel_acesso > 4 Then
        If PrevisaoVendaPrazo.Empresa < g_cfg_empresa_i Or PrevisaoVendaPrazo.Empresa > g_cfg_empresa_f Then
            x_flag = False
        ElseIf PrevisaoVendaPrazo.Data < g_cfg_data_i Or PrevisaoVendaPrazo.Data > g_cfg_data_f Then
            x_flag = False
        End If
    End If
    If x_flag Then
        cmd_alterar.Enabled = True
        cmd_excluir.Enabled = True
    Else
        cmd_alterar.Enabled = False
        cmd_excluir.Enabled = False
    End If
End Sub
Function VerificaLiberacaoDigitacao2() As Boolean
    VerificaLiberacaoDigitacao2 = False
    If g_nivel_acesso <= 4 Then
        VerificaLiberacaoDigitacao2 = True
        Exit Function
    End If
    If msk_data.Text < g_cfg_data_i Or msk_data.Text > g_cfg_data_f Then
        MsgBox "A data do movimento deve estar entre " & Format(g_cfg_data_i, "dd/mm/yyyy") & " a " & Format(g_cfg_data_f, "dd/mm/yyyy") & ".", vbInformation, "Digitação Não Autorizada!"
        msk_data.SetFocus
    Else
        VerificaLiberacaoDigitacao2 = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    'consulta_movimento_composicao_caixa.Show 1
    'If Len(g_string) > 0 Then
    '    lData = RetiraGString(1)
    '    lPeriodo = RetiraGString(2)
    '    lIlha = RetiraGString(3)
    '    lTipoMovimento = RetiraGString(4)
    '    lCodigoComposicao = RetiraGString(5)
    '    If PrevisaoVendaPrazo.LocalizarCodigo(g_empresa, lData, lIlha, lPeriodo, lTipoMovimento, lCodigoComposicao) Then
    '        AtualTela
    '        AtualizaGrid
    '    End If
    'End If
End Sub
Private Sub cmd_primeiro_Click()
    If PrevisaoVendaPrazo.LocalizarPrimeiro Then
        AtualTela
        AtualizaGrid
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If PrevisaoVendaPrazo.LocalizarProximo Then
        AtualTela
        AtualizaGrid
    Else
        MsgBox "Fim de Arquivo.", vbInformation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If PrevisaoVendaPrazo.LocalizarUltimo(g_empresa) Then
        AtualTela
        AtualizaGrid
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub ProximaCelula()
    If fgd_composicao_caixa.Col < fgd_composicao_caixa.Cols - 2 Then
        fgd_composicao_caixa.Col = LastCol + 1
    Else
        fgd_composicao_caixa.Col = 0
        If fgd_composicao_caixa.Row >= fgd_composicao_caixa.Rows - 1 Then
            fgd_composicao_caixa.Rows = fgd_composicao_caixa.Rows + 1
        End If
        fgd_composicao_caixa.Row = fgd_composicao_caixa.Row + 1
    End If
    fgd_composicao_caixa.SetFocus
End Sub
Private Sub CalculaSaldo()
    Dim xSaldo As Currency
    Dim i As Integer
    xSaldo = PrevisaoVendaPrazo.SaldoAnterior(g_empresa, dtcbo_combustivel.BoundText, CDate(msk_data.Text), CDate("00:00:00"))
    With fgd_composicao_caixa
        For i = 1 To (.Rows - 1)
            If Len(.TextMatrix(i, 0)) > 0 Then
                xSaldo = xSaldo + fValidaValor(.TextMatrix(i, 1)) - fValidaValor(.TextMatrix(i, 4))
                .TextMatrix(i, 5) = Format(xSaldo, "###,###,##0.00")
            End If
        Next
    End With
End Sub
Private Sub dtcbo_combustivel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        fgd_composicao_caixa.SetFocus
    End If
End Sub
Private Sub fgd_composicao_caixa_Click()
    ' Quando clicar uma vez
    ' atribui o valor selecionado
    lMarcaCelula = True
    If fgd_composicao_caixa.Col < 5 Then
        LastRow = fgd_composicao_caixa.Row
        LastCol = fgd_composicao_caixa.Col
        txt_celula.Visible = False
    End If
    'AtribuiValorCelula
End Sub
Private Sub fgd_composicao_caixa_DblClick()
    'editar ao clicar duas vezes
    lMarcaCelula = True
    If fgd_composicao_caixa.Col < 5 Then
        '0 - Código da Composicao do Caixa
        '2 - Valor
        LastRow = fgd_composicao_caixa.Row
        LastCol = fgd_composicao_caixa.Col
        txt_celula.Visible = False
        ExibirCelula
    End If
End Sub
Private Sub fgd_composicao_caixa_KeyPress(KeyAscii As Integer)
    lMarcaCelula = True
    Select Case KeyAscii
    ' Editar ao teclar ENTER
    Case vbKeyReturn
        KeyAscii = 0
        If fgd_composicao_caixa.Col < 5 Then
            ExibirCelula
        End If
    ' Cancelar ao pressionar ESC
    Case vbKeyEscape
        KeyAscii = 0
        AtribuiValorCelula
    ' Editar ao pressinar qualquer tecla
    Case 32 To 255
        lMarcaCelula = False
        If fgd_composicao_caixa.Col < 5 Then
            ExibirCelula
            With txt_celula
                If .Visible Then
                    .Text = Chr$(KeyAscii)
                    .SelStart = Len(.Text) + 1
                End If
            End With
        End If
    End Select
End Sub
Private Sub fgd_composicao_caixa_Scroll()
    ' Ver se a coluna esta visivel
    ' entao ocultar os controles
    '
    If fgd_composicao_caixa.ColIsVisible(LastCol) = False Then
        txt_celula.Visible = False
        Exit Sub
    End If
    If fgd_composicao_caixa.RowIsVisible(LastRow) = False Then
        txt_celula.Visible = False
        Exit Sub
    End If
    ' ver se estava visivel antes de ocultar
    ' e posicionar na mesma celula
    If ControlVisible Then
        ExibirCelula
    End If
End Sub
Private Sub Form_Activate()
    If g_empresa <> lEmpresa Then
        lFlagMovimento = 0
    End If
    If lFlagMovimento = 0 Then
        lOpcao = 0
        lEmpresa = g_empresa
        DesativaBotoes
        If PrevisaoVendaPrazo.LocalizarUltimo(g_empresa) Then
            AtualTela
            AtualizaGrid
            AtivaBotoes
        Else
            LimpaGrid
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        If cmd_novo.Enabled Then
            cmd_novo.SetFocus
        End If
    Else
        lFlagMovimento = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_Deactivate()
    lFlagMovimento = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And lOpcao = 0 Then
        KeyCode = 0
        cmd_novo_Click
    ElseIf KeyCode = vbKeyF3 And lOpcao = 0 Then
        KeyCode = 0
        cmd_alterar_Click
    ElseIf KeyCode = vbKeyF4 And Shift = 0 And lOpcao = 0 Then
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
    adodc_combustivel.ConnectionString = gConnectionString
    adodc_combustivel.RecordSource = "SELECT Codigo, Nome FROM Combustivel WHERE Empresa = " & g_empresa & " ORDER BY Nome"
    adodc_combustivel.Refresh
    lData = "01/01/1900"
    lTipoCombustivel = ""
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
        dtcbo_combustivel.SetFocus
    End If
End Sub
Private Sub txt_celula_GotFocus()
    With txt_celula
        If LastCol = 0 Then
            .MaxLength = 10
            If .Text = "" Then
                .Text = Format(Time, "hhmmss")
            End If
        ElseIf LastCol = 2 Then
            .MaxLength = 10
            If .Text = "" Then
                .Text = Format(PrevisaoVendaPrazo.MediaAnterior(g_empresa, dtcbo_combustivel.BoundText, CDate(msk_data.Text), CDate("00:00:00")), "###,###,##0.00")
            End If
            .Text = fValidaValor(.Text)
        ElseIf LastCol <> 0 Then
            .MaxLength = 10
            .Text = fValidaValor(.Text)
        End If
        If lMarcaCelula Then
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
    End With
End Sub
Private Sub txt_celula_KeyPress(KeyAscii As Integer)
    ' ao pressionar ENTER aceitar a entrada de dados
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        fgd_composicao_caixa.SetFocus
    ' ESC, cancela a edição
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        txt_celula.Visible = False
        ControlVisible = False
    End If
    If LastCol = 0 Then
        Call ValidaInteiro(KeyAscii)
    ElseIf LastCol = 2 Then
        If KeyAscii = 46 Then
            KeyAscii = 44
        End If
        Call ValidaValor(KeyAscii)
    End If
End Sub
Private Sub txt_celula_LostFocus()
    'Código do Produto
    If LastCol = 0 Then
        'If Not IsNumeric(txt_celula.Text) Then
        '    MsgBox "Informe o código do serviço.", vbInformation, "Validação Incorreta!"
        '    Exit Sub
        'End If
        If Len(txt_celula.Text) = 6 Then
            txt_celula.Text = Mid(txt_celula.Text, 1, 2) & ":" & Mid(txt_celula.Text, 3, 2) & ":" & Mid(txt_celula.Text, 5, 2)
        ElseIf Len(txt_celula.Text) = 5 Then
            txt_celula.Text = "0" & Mid(txt_celula.Text, 1, 1) & ":" & Mid(txt_celula.Text, 2, 2) & ":" & Mid(txt_celula.Text, 4, 2)
        ElseIf Len(txt_celula.Text) = 4 Then
            txt_celula.Text = Mid(txt_celula.Text, 1, 2) & ":" & Mid(txt_celula.Text, 3, 2) & ":00"
        ElseIf Len(txt_celula.Text) = 3 Then
            txt_celula.Text = "0" & Mid(txt_celula.Text, 1, 1) & ":" & Mid(txt_celula.Text, 2, 2) & ":00"
        ElseIf Len(txt_celula.Text) = 2 Then
            txt_celula.Text = Mid(txt_celula.Text, 1, 2) & ":00:00"
        ElseIf Len(txt_celula.Text) = 1 Then
            txt_celula.Text = "0" & Mid(txt_celula.Text, 1, 1) & ":00:00"
        End If
        If IsDate(txt_celula.Text) Then
            AtribuiValorCelula
            'fgd_composicao_caixa.Col = 1
            'fgd_composicao_caixa.SetFocus
            'LastCol = 1
        ElseIf txt_celula.Text = "" Then
            AtribuiValorCelula
        Else
            AtribuiValorCelula
            fgd_composicao_caixa.TextMatrix(LastRow, 0) = ""
            fgd_composicao_caixa.TextMatrix(LastRow, 1) = ""
            fgd_composicao_caixa.TextMatrix(LastRow, 2) = ""
            fgd_composicao_caixa.TextMatrix(LastRow, 3) = ""
            fgd_composicao_caixa.TextMatrix(LastRow, 4) = ""
            fgd_composicao_caixa.TextMatrix(LastRow, 5) = ""
            CalculaSaldo
            cmd_ok.SetFocus
            Exit Sub
        End If
    ElseIf LastCol >= 1 Then
        If fValidaValor(txt_celula.Text) > 0 Then
            txt_celula.Text = Format(txt_celula.Text, "##,###,##0.00")
        Else
            txt_celula.Text = "0,00"
        End If
        AtribuiValorCelula
        CalculaSaldo
    End If
    'If LastCol <> 0 Then
        ProximaCelula
    'End If
End Sub
