VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form MovimentoMedicaoCombustivelRegua 
   Caption         =   "Medição dos Combustíveis pela Régua (Abertura)"
   ClientHeight    =   4080
   ClientLeft      =   2040
   ClientTop       =   1875
   ClientWidth     =   11565
   Icon            =   "MovimentoMedicaoCombustivelRegua.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "MovimentoMedicaoCombustivelRegua.frx":030A
   ScaleHeight     =   4080
   ScaleWidth      =   11565
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2760
      Picture         =   "MovimentoMedicaoCombustivelRegua.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   3120
      Width           =   795
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   60
      Picture         =   "MovimentoMedicaoCombustivelRegua.frx":1BC2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cria um novo registro."
      Top             =   3120
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   960
      Picture         =   "MovimentoMedicaoCombustivelRegua.frx":3254
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Altera o registro atual."
      Top             =   3120
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1860
      Picture         =   "MovimentoMedicaoCombustivelRegua.frx":474E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Exclui o registro atual."
      Top             =   3120
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3660
      Picture         =   "MovimentoMedicaoCombustivelRegua.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3120
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   2955
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11415
      Begin VB.TextBox txtData 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txt_celula 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   6300
         TabIndex        =   17
         Top             =   2160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmd_transfere_dados_lmc 
         Caption         =   "&Transfere p/ LMC"
         Height          =   735
         Left            =   9900
         Picture         =   "MovimentoMedicaoCombustivelRegua.frx":7472
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Transfere as entradas de combustíveis para o LMC."
         Top             =   120
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid fgd_dados 
         Height          =   2235
         Left            =   60
         TabIndex        =   16
         Top             =   660
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   3942
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColorBkg    =   12632256
         AllowUserResizing=   1
      End
      Begin VB.Label Label5 
         Caption         =   "&Data da Medição"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   9300
      TabIndex        =   10
      Top             =   3000
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "MovimentoMedicaoCombustivelRegua.frx":8864
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "MovimentoMedicaoCombustivelRegua.frx":9D5E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "MovimentoMedicaoCombustivelRegua.frx":B258
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "MovimentoMedicaoCombustivelRegua.frx":C6CA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   9840
      Picture         =   "MovimentoMedicaoCombustivelRegua.frx":DC4C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Confirma o registro atual."
      Top             =   3120
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   10680
      Picture         =   "MovimentoMedicaoCombustivelRegua.frx":EF26
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancela o registro atual."
      Top             =   3120
      Width           =   795
   End
End
Attribute VB_Name = "MovimentoMedicaoCombustivelRegua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagMovimento As Integer
Dim lOpcao As Integer
Dim lEmpresa As Integer
Dim lData As Date
Dim lTipoCombustivel As String
Dim lNumeroTanque As Integer
'Dim tbl_med_comb

Private ControlVisible As Boolean     ' Se o controle esta visivel ou nao
Private LastRow As Long               ' Ultima linha em que se editou
Private LastCol As Long               ' ultima coluna em que se editou

Private rsAuxiliar As New adodb.Recordset
Private Combustivel As New cCombustivel
Private ConvMedicaoComb As New cConvMedicaoComb
Private LivroLMC As New cLivroLMC
Private MedicaoCombustivel As New cMedicaoCombustivel
Private TanqueCombustivel As New cTanqueCombustivel


Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_excluir.Enabled = True
    cmd_alterar.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    cmd_transfere_dados_lmc.Visible = False
    frm_move.Visible = True
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
        fgd_dados.TextMatrix(LastRow, LastCol) = Texto
        'If Val(fgd_dados.Text) < 6 Then
        '     fgd_dados.CellForeColor = vbRed
        'Else
        '     fgd_dados.CellForeColor = vbBlue
        'End If
      Case Else
        'If LastRow = 0 And LastCol = 0 Then
            LastRow = fgd_dados.Row
            LastCol = fgd_dados.Col
        'End If
      
        Texto = txt_celula.Text
        fgd_dados.TextMatrix(LastRow, LastCol) = Texto
    End Select
End Sub
Private Sub AtualizaGrid()
    Dim i As Integer
    Dim xSQL As String
    
    LimpaGrid
    i = 0
    fgd_dados.Visible = False
    
    xSQL = ""
    xSQL = xSQL & "   SELECT Nome, [Numero do Tanque], [Tipo de Combustivel], [Medida da Regua], Quantidade, [Observacao 1], [Observacao 2], [Observacao 3], [Desconto Dia Anterior]"
    xSQL = xSQL & "     FROM " & MedicaoCombustivel.NomeTabela & ", Combustivel"
    xSQL = xSQL & "    WHERE " & MedicaoCombustivel.NomeTabela & ".Empresa = " & g_empresa
    xSQL = xSQL & "      AND " & MedicaoCombustivel.NomeTabela & ".Data = " & preparaData(CDate(txtData.Text))
    xSQL = xSQL & "      AND Combustivel.Empresa = " & g_empresa
    xSQL = xSQL & "      AND Combustivel.Codigo = " & MedicaoCombustivel.NomeTabela & ".[Tipo de Combustivel]"
    xSQL = xSQL & " ORDER BY [Numero do Tanque], [Tipo de Combustivel]"
    Set rsAuxiliar = New adodb.Recordset
    Set rsAuxiliar = Conectar.RsConexao(xSQL)
    If rsAuxiliar.RecordCount > 0 Then
        If Not rsAuxiliar.EOF Then
            Do Until rsAuxiliar.EOF
                i = i + 1
                fgd_dados.Rows = fgd_dados.Rows + 1
                fgd_dados.Row = i
                fgd_dados.Col = 0
                fgd_dados.Text = rsAuxiliar("Numero do Tanque").Value
                fgd_dados.Col = 1
                fgd_dados.Text = rsAuxiliar("Nome").Value
                fgd_dados.Col = 2
                fgd_dados.Text = Format(rsAuxiliar("Medida da Regua").Value, "###,##0.00")
                fgd_dados.Col = 3
                fgd_dados.Text = Format(rsAuxiliar("Quantidade").Value, "###,##0.00")
                fgd_dados.Col = 4
                fgd_dados.Text = Format(rsAuxiliar("Desconto Dia Anterior").Value, "###,##0.00")
                fgd_dados.Col = 5
                fgd_dados.Text = rsAuxiliar("Observacao 1").Value
                fgd_dados.Col = 6
                fgd_dados.Text = rsAuxiliar("Observacao 2").Value
                fgd_dados.Col = 7
                fgd_dados.Text = rsAuxiliar("Observacao 3").Value
                lTipoCombustivel = rsAuxiliar("Tipo de Combustivel").Value
                lNumeroTanque = rsAuxiliar("Numero do Tanque").Value
                rsAuxiliar.MoveNext
            Loop
        End If
        fgd_dados.Row = 1
        fgd_dados.Col = 3
    End If
    fgd_dados.Visible = True
    frm_dados.Enabled = False
    rsAuxiliar.Close
    Set rsAuxiliar = Nothing
End Sub
Private Sub AtualTabe()
    Dim i As Integer
    Dim xString As String
    
    For i = 1 To (fgd_dados.Rows - 1)
        If fgd_dados.TextMatrix(i, 0) <> "" Then
            If Combustivel.LocalizarNome(g_empresa, fgd_dados.TextMatrix(i, 1)) Then
                MedicaoCombustivel.Empresa = g_empresa
                MedicaoCombustivel.Data = Format(txtData.Text, "dd/mm/yyyy")
                MedicaoCombustivel.NumeroTanque = Val(fgd_dados.TextMatrix(i, 0))
                MedicaoCombustivel.TipoCombustivel = Combustivel.Codigo
                MedicaoCombustivel.Quantidade = CCur(fgd_dados.TextMatrix(i, 3))
                MedicaoCombustivel.Observacao1 = fgd_dados.TextMatrix(i, 5)
                MedicaoCombustivel.Observacao2 = fgd_dados.TextMatrix(i, 6)
                MedicaoCombustivel.Observacao3 = fgd_dados.TextMatrix(i, 7)
                MedicaoCombustivel.DescontoDiaAnterior = CCur(fgd_dados.TextMatrix(i, 4))
                MedicaoCombustivel.MedidaRegua = Val(fgd_dados.TextMatrix(i, 2))
                xString = ""
                If lOpcao = 2 Then
                    xString = "Para: "
                End If
                xString = xString & "Tanque:" & MedicaoCombustivel.NumeroTanque
                xString = xString & " Comb:" & MedicaoCombustivel.TipoCombustivel
                xString = xString & " Qtd:" & MedicaoCombustivel.Quantidade
                xString = xString & " Desc:" & MedicaoCombustivel.DescontoDiaAnterior
                Call GravaAuditoria(1, Me.name, 10, xString)
                If MedicaoCombustivel.Incluir Then
                    lData = CDate(txtData.Text)
                    lTipoCombustivel = Combustivel.Codigo
                    lNumeroTanque = MedicaoCombustivel.NumeroTanque
                Else
                    MsgBox "Registro não foi gravado!", vbInformation, "Erro de Integridade"
                End If
            Else
                MsgBox "Combustível não Cadastrado", vbInformation, "Erro de Integridade"
            End If
        End If
    Next
End Sub
Private Sub AtualTela()
    lData = MedicaoCombustivel.Data
    txtData.Text = Format(MedicaoCombustivel.Data, "dd/mm/yyyy")
    frm_dados.Enabled = True
    Call VerificaLiberacaoLMC("**", lData)
End Sub
Private Sub AutomatizaGridInclusao()
    Dim i As Integer
    Dim xSQL As String
    
    For i = 1 To (fgd_dados.Rows - 2)
        If fgd_dados.TextMatrix(i, 0) <> "" Then
            Exit Sub
        End If
    Next
    i = 0
    xSQL = ""
    xSQL = xSQL & "   SELECT [Numero do Tanque], Nome"
    xSQL = xSQL & "     FROM Tanque_Combustivel, Combustivel "
    xSQL = xSQL & "    WHERE Tanque_Combustivel.Empresa = " & g_empresa
    xSQL = xSQL & "      AND Combustivel.Empresa = " & g_empresa
    xSQL = xSQL & "      AND Combustivel.Codigo = Tanque_Combustivel.[Tipo de Combustivel]"
    xSQL = xSQL & " ORDER BY [Numero do Tanque]"
    Set rsAuxiliar = New adodb.Recordset
    Set rsAuxiliar = Conectar.RsConexao(xSQL)
    If Not rsAuxiliar.EOF Then
        Do Until rsAuxiliar.EOF
            i = i + 1
            fgd_dados.Rows = fgd_dados.Rows + 1
            fgd_dados.Row = i
            fgd_dados.Col = 0
            fgd_dados.Text = Format(rsAuxiliar("Numero do Tanque").Value, "#,##0")
            fgd_dados.Col = 1
            fgd_dados.Text = rsAuxiliar("Nome").Value
            fgd_dados.Col = 2
            fgd_dados.Text = "0"
            fgd_dados.Col = 3
            fgd_dados.Text = "0"
            fgd_dados.Col = 4
            fgd_dados.Text = "0"
            fgd_dados.Col = 5
            fgd_dados.Text = ""
            fgd_dados.Col = 6
            fgd_dados.Text = ""
            fgd_dados.Col = 7
            fgd_dados.Text = ""
            rsAuxiliar.MoveNext
        Loop
    End If
    rsAuxiliar.Close
    Set rsAuxiliar = Nothing
    
    txt_celula.Visible = False
    fgd_dados.Row = 1
    fgd_dados.Col = 2
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_excluir.Enabled = False
    cmd_alterar.Enabled = False
    cmd_pesquisa.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    If g_nome_usuario = "L.M.C." Then
        cmd_transfere_dados_lmc.Visible = True
    End If
End Sub
Private Sub ExibirCelula()
    Static OK As Boolean
    '
    ' Se for celula fixa , sair
    If fgd_dados.Col <= fgd_dados.FixedCols - 1 Or fgd_dados.Row <= fgd_dados.FixedRows - 1 Then
       Exit Sub
    End If
    
    If OK Then Exit Sub
    OK = True
    '
    txt_celula.Visible = False
    '
    LastRow = fgd_dados.Row
    LastCol = fgd_dados.Col
    '
    Select Case LastCol
        Case Else
        txt_celula.Move fgd_dados.CellLeft - Screen.TwipsPerPixelX + 60, fgd_dados.CellTop + 650 - Screen.TwipsPerPixelY, fgd_dados.CellWidth + Screen.TwipsPerPixelX * 2, fgd_dados.CellHeight + Screen.TwipsPerPixelY * 2
        txt_celula.Text = fgd_dados.Text
        'If Len(fgd_dados.Text) = 0 Then
        '   If LastRow > 1 Then
        '       txt_celula.Text = fgd_dados.TextMatrix(LastRow - 1, LastCol)
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
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Combustivel = Nothing
    Set ConvMedicaoComb = Nothing
    Set LivroLMC = Nothing
    Set MedicaoCombustivel = Nothing
    Set TanqueCombustivel = Nothing
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    If IsDate(txtData.Text) Then
        txtData.Text = Format(CDate(txtData.Text) + 1, "dd/mm/yyyy")
    Else
        txtData.Text = Format(g_data_def, "dd/mm/yyyy")
    End If
End Sub
Private Sub cmd_alterar_Click()
    Call GravaAuditoria(1, Me.name, 3, "")
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frm_dados.Enabled = True
    fgd_dados.Row = 1
    fgd_dados.Col = 2
    fgd_dados.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    Call GravaAuditoria(1, Me.name, 13, "")
    If MedicaoCombustivel.LocalizarAnterior(g_empresa, lData) Then
        AtualTela
        AtualizaGrid
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    Call GravaAuditoria(1, Me.name, 9, "")
    LimpaTela
    If MedicaoCombustivel.LocalizarCodigo(g_empresa, lData, lNumeroTanque) Then
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
Private Sub LimpaGrid()
    Dim i As Integer
    Dim xSQL As String
    fgd_dados.WordWrap = True
    fgd_dados.Rows = 1
    fgd_dados.RowHeight(0) = 500
    fgd_dados.Row = 0
    i = 0
    fgd_dados.Col = i
    fgd_dados.Text = "Tanque"
    fgd_dados.ColWidth(i) = 500
    fgd_dados.ColAlignment(i) = 7
    i = i + 1
    fgd_dados.Col = i
    fgd_dados.Text = "Combustível"
    fgd_dados.ColWidth(i) = 1500
    fgd_dados.ColAlignment(i) = 1
    i = i + 1
    fgd_dados.Col = i
    fgd_dados.Text = "Medida da Regua  "
    fgd_dados.ColWidth(i) = 900
    fgd_dados.ColAlignment(i) = 7
    i = i + 1
    fgd_dados.Col = i
    fgd_dados.Text = "Quantidade em Litros  "
    fgd_dados.ColWidth(i) = 900
    fgd_dados.ColAlignment(i) = 7
    i = i + 1
    fgd_dados.Col = i
    fgd_dados.Text = "Desconto do Dia Anterior"
    fgd_dados.ColWidth(i) = 1000
    fgd_dados.ColAlignment(i) = 7
    i = i + 1
    fgd_dados.Col = i
    fgd_dados.Text = "Observação Linha 1"
    fgd_dados.ColWidth(i) = 2900
    fgd_dados.ColAlignment(i) = 1
    i = i + 1
    fgd_dados.Col = i
    fgd_dados.Text = "Observação Linha 2"
    fgd_dados.ColWidth(i) = 2900
    fgd_dados.ColAlignment(i) = 1
    i = i + 1
    fgd_dados.Col = i
    fgd_dados.Text = "Observação Linha 3"
    fgd_dados.ColWidth(i) = 2900
    fgd_dados.ColAlignment(i) = 1
End Sub
Private Sub LimpaTela()
    LimpaGrid
End Sub
Private Sub cmd_excluir_Click()
    Call GravaAuditoria(1, Me.name, 4, "")
    If IsDate(txtData.Text) Then
        If (MsgBox("Deseja excluir estes registros?", 4 + 32 + 256, "Exclusão de Registros!")) = 6 Then
            Call GravaAuditoria(1, Me.name, 10, "Data:" & txtData.Text)
            If MedicaoCombustivel.ExcluirRegistros(g_empresa, CDate(txtData.Text)) Then
                LimpaTela
                If MedicaoCombustivel.LocalizarUltimo(g_empresa) Then
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
'    zzClonaMedicao
'    zzSomaTanque
'    Exit Sub
    
    Call GravaAuditoria(1, Me.name, 2, "")
    LimpaTela
    Inclui
    frm_dados.Enabled = True
    AutomatizaGridInclusao
    fgd_dados.SetFocus
End Sub
Private Sub cmd_ok_Click()
    'On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            Call GravaAuditoria(1, Me.name, 10, "Data:" & txtData.Text)
            AtualTabe
            cmd_novo.SetFocus
        ElseIf lOpcao = 2 Then
            Call GravaAuditoria(1, Me.name, 10, "De: Data:" & Format(lData, "dd/mm/yyyy"))
            Call GravaAuditoria(1, Me.name, 10, "Para: Data:" & txtData.Text)
            Call MedicaoCombustivel.ExcluirRegistros(g_empresa, lData)
            AtualTabe
        End If
        lOpcao = 0
        If MedicaoCombustivel.LocalizarCodigo(g_empresa, lData, lNumeroTanque) Then
            AtualizaGrid
        Else
            MsgBox "Registro não Encontrado", vbInformation, "Erro de Integridade"
        End If
        cmd_novo.SetFocus
    End If
    Exit Sub
'FileError:
    MsgBox Error
    Exit Sub
End Sub
Private Sub ProximaCelula()
    If fgd_dados.Col < fgd_dados.Cols - 1 And fgd_dados.Col <> 2 Then
        fgd_dados.Col = LastCol + 1
    Else
        fgd_dados.Col = 2
        If fgd_dados.Row >= fgd_dados.Rows - 1 Then
            cmd_ok.SetFocus
            Exit Sub
        End If
        fgd_dados.Row = fgd_dados.Row + 1
    End If
    fgd_dados.SetFocus
End Sub
Private Sub TransfereDadosLMC()
    Dim x_data As Date
    Dim xSQL As String
    
    On Error GoTo FileError
    
    x_data = CDate("01/01/1900")
    'Busca ultima data com movimento
    If MedicaoCombustivel.LocalizarUltimo(g_empresa) Then
        x_data = MedicaoCombustivel.Data
    End If
    x_data = x_data + 1
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será transferido a medição de combustível apartir da data " & x_data & "." & Chr(10) & Chr(10) & "Deseja realmente fazer esta transferência?", vbYesNo + 256, "Transfere a Medição de Combustível Para o L.M.C.!")) = vbNo Then
        Exit Sub
    End If
    Call GravaAuditoria(1, Me.name, 10, "Empresa:" & g_empresa & " A Partir de:" & x_data)
    
    'Seleciona Registros a Serem Transferidos
    xSQL = ""
    xSQL = xSQL & "   SELECT Data, [Numero do Tanque], [Tipo de Combustivel], Quantidade,"
    xSQL = xSQL & "          [Observacao 1], [Observacao 2], [Observacao 3], [Desconto Dia Anterior]"
    xSQL = xSQL & "     FROM MedicaoCombustivel"
    xSQL = xSQL & "    WHERE Empresa = " & g_empresa
    xSQL = xSQL & "      AND Data >= " & preparaData(x_data)
    xSQL = xSQL & " ORDER BY Data, [Numero do Tanque]"
    Set rsAuxiliar = New adodb.Recordset
    Set rsAuxiliar = Conectar.RsConexao(xSQL)
    
    'Transfere Dados para o LMC
    If Not rsAuxiliar.EOF Then
        Do Until rsAuxiliar.EOF
            MedicaoCombustivel.Empresa = g_empresa
            MedicaoCombustivel.Data = rsAuxiliar("Data").Value
            MedicaoCombustivel.NumeroTanque = rsAuxiliar("Numero do Tanque").Value
            MedicaoCombustivel.TipoCombustivel = rsAuxiliar("Tipo de Combustivel").Value
            MedicaoCombustivel.Quantidade = rsAuxiliar("Quantidade").Value
            MedicaoCombustivel.Observacao1 = rsAuxiliar("Observacao 1").Value
            MedicaoCombustivel.Observacao2 = rsAuxiliar("Observacao 2").Value
            MedicaoCombustivel.Observacao3 = rsAuxiliar("Observacao 3").Value
            MedicaoCombustivel.DescontoDiaAnterior = rsAuxiliar("Desconto Dia Anterior").Value
            If Not MedicaoCombustivel.Incluir Then
                MsgBox "Registro não foi gravado!", vbInformation, "Erro de Integridade"
            End If
            rsAuxiliar.MoveNext
        Loop
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com a medição de combustível transferida para o L.M.C.", vbInformation, "Transferência Concluida!"
    Else
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Não tem medição de combustível à ser transferida para o L.M.C.", vbInformation, "Transferência Não Concluida!"
    End If
    rsAuxiliar.Close
    Set rsAuxiliar = Nothing
    If MedicaoCombustivel.LocalizarUltimo(g_empresa) Then
        lData = MedicaoCombustivel.Data
    End If
    Exit Sub

FileError:
    MsgBox Error
    Resume Next

End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(txtData.Text) Then
        MsgBox "Informe a Data da Medição.", vbInformation, "Atenção!"
        txtData.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Function VerificaLiberacaoLMC(ByVal pTipoCombustivel As String, ByVal pData As Date) As Boolean
    If g_nome_usuario = "L.M.C." Then
        VerificaLiberacaoLMC = False
        If LivroLMC.LocalizarCombustivelConcluido(g_empresa, pTipoCombustivel, CDate(pData - 1)) = "NAO" Then
            VerificaLiberacaoLMC = True
            cmd_alterar.Enabled = True
            cmd_excluir.Enabled = True
        Else
            cmd_alterar.Enabled = False
            cmd_excluir.Enabled = False
        End If
    Else
        VerificaLiberacaoLMC = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    Call GravaAuditoria(1, Me.name, 5, "")
    consulta_medicao_combustiveis.Show 1
    If Len(g_string) > 0 Then
        lData = RetiraGString(1)
        lNumeroTanque = RetiraGString(2)
        If MedicaoCombustivel.LocalizarCodigo(g_empresa, lData, lNumeroTanque) Then
            AtualTela
            AtualizaGrid
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    Call GravaAuditoria(1, Me.name, 12, "")
    If MedicaoCombustivel.LocalizarPrimeiro Then
        AtualTela
        AtualizaGrid
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    Call GravaAuditoria(1, Me.name, 14, "")
    If MedicaoCombustivel.LocalizarProximo(g_empresa, lData) Then
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
Private Sub cmd_transfere_dados_lmc_Click()
    Call GravaAuditoria(1, Me.name, 23, "Transferencia Para LMC")
    If MedicaoCombustivel.TransfereDadosLMC(g_empresa, True) Then
        Call GravaAuditoria(1, Me.name, 10, "Empresa:" & g_empresa & " A Partir de:" & MedicaoCombustivel.UltimaData(g_empresa))
        If MedicaoCombustivel.TransfereDadosLMC(g_empresa, False) Then
            MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com a medição de combustível transferida para o L.M.C.", vbInformation, "Transferência Concluida!"
        Else
            MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Não tem medição de combustível à ser transferida para o L.M.C.", vbInformation, "Transferência Não Concluida!"
        End If
    Else
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Não tem medição de combustível à ser transferida para o L.M.C.", vbInformation, "Transferência Não Concluida!"
    End If
    'TransfereDadosLMC
    cmd_cancelar_Click
    cmd_ultimo_Click
End Sub
Private Sub cmd_ultimo_Click()
    Call GravaAuditoria(1, Me.name, 15, "")
    If MedicaoCombustivel.LocalizarUltimo(g_empresa) Then
        AtualTela
        AtualizaGrid
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub fgd_dados_Click()
    ' Quando clicar uma vez
    ' atribui o valor selecionado
    If fgd_dados.Col >= 1 And fgd_dados.Col <> 3 Then
        LastRow = fgd_dados.Row
        LastCol = fgd_dados.Col
        txt_celula.Visible = False
    End If
    'AtribuiValorCelula
End Sub
Private Sub fgd_dados_DblClick()
    'editar ao clicar duas vezes
    If fgd_dados.Col >= 2 And fgd_dados.Col <> 3 Then
        '0 - Código da Composicao do Caixa
        '2 - Valor
        LastRow = fgd_dados.Row
        LastCol = fgd_dados.Col
        txt_celula.Visible = False
        ExibirCelula
    End If
End Sub
Private Sub fgd_dados_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    ' Editar ao teclar ENTER
    Case vbKeyReturn
        KeyAscii = 0
        If fgd_dados.Col >= 2 And fgd_dados.Col <> 3 Then
            ExibirCelula
        End If
    ' Cancelar ao pressionar ESC
    Case vbKeyEscape
        KeyAscii = 0
        AtribuiValorCelula
    ' Editar ao pressinar qualquer tecla
    Case 32 To 255
        If fgd_dados.Col >= 2 And fgd_dados.Col <> 3 Then
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
Private Sub fgd_dados_Scroll()
    ' Ver se a coluna esta visivel
    ' entao ocultar os controles
    '
    If fgd_dados.ColIsVisible(LastCol) = False Then
        txt_celula.Visible = False
        Exit Sub
    End If
    If fgd_dados.RowIsVisible(LastRow) = False Then
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
    Call GravaAuditoria(1, Me.name, 1, "")
    If g_empresa <> lEmpresa Then
        lFlagMovimento = 0
    End If
    If lFlagMovimento = 0 Then
        lOpcao = 0
        lEmpresa = g_empresa
        DesativaBotoes
        If MedicaoCombustivel.LocalizarUltimo(g_empresa) Then
            AtualTela
            AtualizaGrid
            AtivaBotoes
            Call VerificaLiberacaoLMC("**", lData)
        Else
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
    
    If g_nome_usuario = "L.M.C." Then
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivelLMC"
        Me.Caption = Me.Caption & " - LMC"
    Else
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivel"
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub

Private Sub txt_celula_GotFocus()
    With txt_celula
        .SelStart = Len(.Text)
        If LastCol = 2 Or LastCol = 4 Then
            .MaxLength = 10
        ElseIf LastCol > 4 Then
            .MaxLength = 40
        End If
    End With
End Sub
Private Sub txt_celula_KeyPress(KeyAscii As Integer)
    ' ao pressionar ENTER aceitar a entrada de dados
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        fgd_dados.SetFocus
    ' ESC, cancela a edição
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        txt_celula.Visible = False
        ControlVisible = False
    End If
    'If LastCol = 1 Then
    '    Call ValidaInteiro(KeyAscii)
    'End If
End Sub
Private Sub txt_celula_LostFocus()
    'Código do Produto
    If LastCol = 2 Then
        If fValidaValor(txt_celula.Text) > 0 Then
            txt_celula.Text = Format(txt_celula.Text, "##,###,##0")
        Else
            txt_celula.Text = "0"
        End If
        fgd_dados.TextMatrix(LastRow, 3) = "0,00"
        If TanqueCombustivel.LocalizarCodigo(g_empresa, Val(fgd_dados.TextMatrix(LastRow, 0))) Then
            If ConvMedicaoComb.LocalizarCodigo(g_empresa, Val(txt_celula.Text)) Then
                If TanqueCombustivel.CapacidadeArmazenamento = 10000 Then
                    fgd_dados.TextMatrix(LastRow, 3) = Format(ConvMedicaoComb.MedicaoTanque10, "##,###,##0.00")
                ElseIf TanqueCombustivel.CapacidadeArmazenamento = 15000 Then
                    fgd_dados.TextMatrix(LastRow, 3) = Format(ConvMedicaoComb.MedicaoTanque15, "##,###,##0.00")
                ElseIf TanqueCombustivel.CapacidadeArmazenamento = 20000 Then
                    fgd_dados.TextMatrix(LastRow, 3) = Format(ConvMedicaoComb.MedicaoTanque20, "##,###,##0.00")
                ElseIf TanqueCombustivel.CapacidadeArmazenamento = 30000 Then
                    fgd_dados.TextMatrix(LastRow, 3) = Format(ConvMedicaoComb.MedicaoTanque30, "##,###,##0.00")
                End If
            End If
        End If
        AtribuiValorCelula
    ElseIf LastCol = 3 Then
        If fValidaValor(txt_celula.Text) > 0 Then
            txt_celula.Text = Format(txt_celula.Text, "##,###,##0.00")
        Else
            txt_celula.Text = "0,00"
        End If
        AtribuiValorCelula
    ElseIf LastCol = 4 Then
        If fValidaValor(txt_celula.Text) > 0 Then
            txt_celula.Text = Format(txt_celula.Text, "##,###,##0.00")
        Else
            txt_celula.Text = "0,00"
        End If
        AtribuiValorCelula
    ElseIf LastCol >= 5 Then
        AtribuiValorCelula
    End If
    'If LastCol <> 1 Then
        ProximaCelula
    'End If
End Sub
Private Sub zzClonaMedicao()
    Dim xTanque As Integer
    Dim xData As Date
    
    Exit Sub
    For xTanque = 1 To 6
        If MedicaoCombustivel.LocalizarCodigo(g_empresa, CDate("01/10/2006"), xTanque) Then
            For xData = CDate("02/10/2006") To CDate("01/08/2007")
                MedicaoCombustivel.Data = xData
                If Not MedicaoCombustivel.Incluir Then
                    MsgBox "Não foi possível incluir"
                End If
            Next
        Else
            MsgBox "Movimento inexistente"
        End If
    Next
End Sub
Private Sub zzSomaTanque()
    Dim xTanqueLer As Integer
    Dim xTanqueAlterar As Integer
    Dim xData As Date
    Dim xQuantidade As Currency
    
    xTanqueLer = 4
    xTanqueAlterar = 3
    xData = CDate("01/04/2007")
    Exit Sub
    Do Until xData = CDate("01/07/2008")
        If MedicaoCombustivel.LocalizarCodigo(g_empresa, xData, xTanqueLer) Then
            xQuantidade = MedicaoCombustivel.Quantidade
            If MedicaoCombustivel.LocalizarCodigo(g_empresa, xData, xTanqueAlterar) Then
                MedicaoCombustivel.Quantidade = MedicaoCombustivel.Quantidade + xQuantidade
                If Not MedicaoCombustivel.Alterar(g_empresa, xData, xTanqueAlterar) Then
                    MsgBox "Não foi possível alterar"
                End If
            Else
                MsgBox "Movimento inexistente - Gravacao"
            End If
        Else
            MsgBox "Movimento inexistente - Leitura"
        End If
        xData = xData + 1
    Loop
End Sub
Private Sub txtData_GotFocus()
    txtData.Text = fDesmascaraData(txtData.Text)
    txtData.SelStart = 0
    txtData.SelLength = 2
    txtData.MaxLength = 8
End Sub
Private Sub txtData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        fgd_dados.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtData_LostFocus()
    txtData.MaxLength = 10
    txtData.Text = fMascaraData(txtData.Text)
End Sub

