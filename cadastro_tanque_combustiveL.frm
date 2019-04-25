VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form cadastro_tanque_combustivel 
   Caption         =   "Cadastro de Tanques de Combustíveis"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   Icon            =   "cadastro_tanque_combustiveL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4890
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm_dados 
      Height          =   1395
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6915
      Begin VB.TextBox txt_numero_tanque 
         Height          =   285
         Left            =   2580
         MaxLength       =   2
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin MSAdodcLib.Adodc adodc_combustivel 
         Height          =   330
         Left            =   3600
         Top             =   540
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
         Bindings        =   "cadastro_tanque_combustiveL.frx":030A
         Height          =   315
         Left            =   2580
         TabIndex        =   4
         Top             =   600
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_combustivel"
      End
      Begin MSAdodcLib.Adodc adodc_capacidade_tanque 
         Height          =   330
         Left            =   3600
         Top             =   900
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
         Caption         =   "adodc_capacidade_tanque"
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
      Begin MSDataListLib.DataCombo dtcbo_capacidade_tanque 
         Bindings        =   "cadastro_tanque_combustiveL.frx":032A
         Height          =   315
         Left            =   2580
         TabIndex        =   6
         Top             =   960
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Capacidade de Armazenamento"
         BoundColumn     =   "Capacidade de Armazenamento"
         Text            =   "dtcbo_capacidade_tanque"
      End
      Begin VB.Label Label2 
         Caption         =   "Capacidade de Armazenamento"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label10 
         Caption         =   "Combustível"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "&Número do Tanque"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cadastro_tanque_combustiveL.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cria um novo registro."
      Top             =   3960
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cadastro_tanque_combustiveL.frx":19E2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Altera o registro atual."
      Top             =   3960
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cadastro_tanque_combustiveL.frx":2EDC
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Exclui o registro atual."
      Top             =   3960
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2820
      Picture         =   "cadastro_tanque_combustiveL.frx":456E
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3960
      Width           =   795
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4860
      TabIndex        =   14
      Top             =   3840
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cadastro_tanque_combustiveL.frx":5C00
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cadastro_tanque_combustiveL.frx":70FA
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cadastro_tanque_combustiveL.frx":85F4
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cadastro_tanque_combustiveL.frx":9A66
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6240
      Picture         =   "cadastro_tanque_combustiveL.frx":AFE8
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Cancela o registro atual."
      Top             =   3960
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5340
      Picture         =   "cadastro_tanque_combustiveL.frx":C4E2
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Confirma o registro atual."
      Top             =   3960
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   4260
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483633
   End
End
Attribute VB_Name = "cadastro_tanque_combustivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lNumeroTanque As Integer
Dim lSQl As String
Private rsTabela As New adodb.Recordset
Private TanqueCombustivel As New cTanqueCombustivel
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    cmd_sair.Enabled = True
    MSFlexGrid.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    txt_numero_tanque.Enabled = True
End Sub
Private Sub AtualizaMSFlexGrid()
    Dim i As Integer
    On Error GoTo ErroConsulta
    LimpaMSFlexGrid
    lSQl = "SELECT Tanque_Combustivel.[Numero do Tanque], Combustivel.Nome as Combustível, Tanque_Combustivel.[Capacidade de Armazenamento]"
    lSQl = lSQl & " FROM Tanque_Combustivel, Combustivel"
    lSQl = lSQl & " WHERE Tanque_Combustivel.Empresa = " & g_empresa
    lSQl = lSQl & " AND Combustivel.Empresa = " & g_empresa
    lSQl = lSQl & " AND Combustivel.Codigo = Tanque_Combustivel.[Tipo de Combustivel]"
    lSQl = lSQl & " ORDER BY [Numero do Tanque]"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQl)
    'Verifica movimento
    i = 0
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            MSFlexGrid.Rows = MSFlexGrid.Rows + 1
            i = i + 1
            MSFlexGrid.Row = i
            MSFlexGrid.Col = 0
            MSFlexGrid.Text = rsTabela("Numero do Tanque").Value
            MSFlexGrid.Col = 1
            MSFlexGrid.Text = rsTabela("Combustível").Value
            MSFlexGrid.Col = 2
            MSFlexGrid.Text = Format(rsTabela("Capacidade de Armazenamento").Value, "####,###,##0")
            rsTabela.MoveNext
        Loop
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
    Exit Sub
    
ErroConsulta:
    If Err = 3075 Then
        MsgBox "Condição inválida.", vbExclamation, "Erro de Consulta"
    Else
        MsgBox Error, vbExclamation, "Erro de Consulta"
    End If
    Exit Sub
End Sub
Private Sub AtualTabe()
    If lOpcao = 1 Then
        TanqueCombustivel.Empresa = g_empresa
        TanqueCombustivel.NumeroTanque = Val(txt_numero_tanque.Text)
    End If
    TanqueCombustivel.TipoCombustivel = dtcbo_combustivel.BoundText
    TanqueCombustivel.CapacidadeArmazenamento = CCur(dtcbo_capacidade_tanque.BoundText)
End Sub
Private Sub AtualTela()
    lNumeroTanque = TanqueCombustivel.NumeroTanque
    txt_numero_tanque.Text = Format(TanqueCombustivel.NumeroTanque, "#0")
    dtcbo_combustivel.BoundText = ""
    dtcbo_combustivel.BoundText = TanqueCombustivel.TipoCombustivel
    dtcbo_capacidade_tanque.BoundText = ""
    dtcbo_capacidade_tanque.BoundText = TanqueCombustivel.CapacidadeArmazenamento
    frm_dados.Enabled = False
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_excluir.Enabled = False
    cmd_sair.Enabled = False
    MSFlexGrid.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub Finaliza()
    Set TanqueCombustivel = Nothing
    frm_cadastro.Show
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_numero_tanque.Text = 1
    If TanqueCombustivel.LocalizarUltimo(g_empresa) Then
        txt_numero_tanque.Text = TanqueCombustivel.NumeroTanque + 1
    End If
End Sub
Private Sub cmd_alterar_Click()
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frm_dados.Enabled = True
    txt_numero_tanque.Enabled = False
    dtcbo_combustivel.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If TanqueCombustivel.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If TanqueCombustivel.LocalizarCodigo(g_empresa, lNumeroTanque) Then
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
    MSFlexGrid.Cols = 3
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
    MSFlexGrid.Text = "Número do Tanque"
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Combustível"
    MSFlexGrid.ColWidth(i) = 2000
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Capacidade de Armazenamento"
    MSFlexGrid.ColWidth(i) = 2000
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
End Sub
Private Sub LimpaTela()
    txt_numero_tanque.Text = ""
    dtcbo_combustivel.BoundText = ""
    dtcbo_capacidade_tanque.BoundText = ""
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_numero_tanque.Text) > 0 Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If TanqueCombustivel.Excluir(g_empresa, Val(txt_numero_tanque.Text)) Then
                LimpaTela
                If TanqueCombustivel.LocalizarUltimo(g_empresa) Then
                    AtualTela
                    AtualizaMSFlexGrid
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
Private Sub cmd_novo_Click()
    LimpaTela
    Inclui
    frm_dados.Enabled = True
    dtcbo_combustivel.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If TanqueCombustivel.Incluir Then
                lNumeroTanque = Val(txt_numero_tanque.Text)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not TanqueCombustivel.Alterar(g_empresa, lNumeroTanque) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        If TanqueCombustivel.LocalizarCodigo(g_empresa, lNumeroTanque) Then
            AtualTela
            AtualizaMSFlexGrid
        End If
        lOpcao = 0
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_TanqueCombustivel.Name, "TanqueCombustivelo"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not Val(txt_numero_tanque.Text) > 0 Then
        MsgBox "Informe o número do tanque.", vbInformation, "Atenção!"
        txt_numero_tanque.SetFocus
    ElseIf IsNull(dtcbo_combustivel.SelectedItem) Then
        MsgBox "Escolha o Combustível.", vbInformation, "Atenção!"
        dtcbo_combustivel.SetFocus
    ElseIf IsNull(dtcbo_capacidade_tanque.SelectedItem) Then
        MsgBox "Escolha a capacidade de armazenamento.", vbInformation, "Atenção!"
        dtcbo_capacidade_tanque.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_primeiro_Click()
    If TanqueCombustivel.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If TanqueCombustivel.LocalizarProximo Then
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
    If TanqueCombustivel.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dtcbo_capacidade_tanque_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub dtcbo_combustivel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        dtcbo_capacidade_tanque.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        DesativaBotoes
        If TanqueCombustivel.LocalizarUltimo(g_empresa) Then
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
    Screen.MousePointer = 1
    CentraForm Me
    adodc_combustivel.ConnectionString = gConnectionString
    adodc_combustivel.RecordSource = "SELECT Codigo, Nome FROM Combustivel WHERE Empresa = " & g_empresa & " ORDER BY Nome"
    adodc_combustivel.Refresh
    adodc_capacidade_tanque.ConnectionString = gConnectionString
    adodc_capacidade_tanque.RecordSource = "SELECT [Capacidade de Armazenamento] FROM Capacidade_Tanque"
    adodc_capacidade_tanque.Refresh
    AtualizaMSFlexGrid
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub MSFlexGrid_RowColChange()
    If lOpcao = 0 And MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0) <> "" Then
        lNumeroTanque = Val(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0))
        If TanqueCombustivel.LocalizarCodigo(g_empresa, lNumeroTanque) Then
            AtualTela
        End If
    End If
End Sub
Private Sub txt_numero_tanque_GotFocus()
    txt_numero_tanque.SelStart = 0
    txt_numero_tanque.SelLength = Len(txt_numero_tanque.Text)
End Sub
Private Sub txt_numero_tanque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcbo_combustivel.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_numero_tanque_LostFocus()
    If lOpcao = 1 And txt_numero_tanque.Text <> "" Then
        If TanqueCombustivel.LocalizarCodigo(g_empresa, Val(txt_numero_tanque.Text)) Then
            MsgBox "Já existe tanque de combustível cadastrado com este código." & Chr(10) & Chr(10) & "Mude o código informado.", vbInformation, "Duplicidade de Registro!"
            txt_numero_tanque.SetFocus
            Exit Sub
        End If
    End If
End Sub
