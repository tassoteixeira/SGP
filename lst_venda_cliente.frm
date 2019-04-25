VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form emissao_venda_cliente 
   Caption         =   "Emissão da Venda por Cliente"
   ClientHeight    =   2775
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   7275
   Icon            =   "lst_venda_cliente.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2775
   ScaleWidth      =   7275
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5220
      Picture         =   "lst_venda_cliente.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3240
      Picture         =   "lst_venda_cliente.frx":199C
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Imprime as vendas por cliente."
      Top             =   1800
      Width           =   795
   End
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1260
      Picture         =   "lst_venda_cliente.frx":2FA6
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Visualiza as vendas por cliente."
      Top             =   1800
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7035
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   6420
         Picture         =   "lst_venda_cliente.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2640
         Picture         =   "lst_venda_cliente.frx":599A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2640
         Picture         =   "lst_venda_cliente.frx":6C74
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_cliente 
         Height          =   285
         Left            =   1500
         MaxLength       =   6
         TabIndex        =   5
         Top             =   660
         Width           =   795
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   5280
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data_i 
         Height          =   315
         Left            =   1500
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   315
         Left            =   1500
         TabIndex        =   2
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc adodcCliente 
         Height          =   330
         Left            =   3780
         Top             =   660
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
         Caption         =   "adodcCliente"
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
      Begin MSDataListLib.DataCombo dtcboCliente 
         Bindings        =   "lst_venda_cliente.frx":7F4E
         Height          =   315
         Left            =   2340
         TabIndex        =   6
         Top             =   660
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Razao Social"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboCliente"
      End
      Begin VB.Label Label3 
         Caption         =   "C&liente"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1140
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1275
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   660
      Top             =   2100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport cr_relato 
      Left            =   120
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\VB5\SGP\Data\SQL_lst_venda_cliente.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
End
Attribute VB_Name = "emissao_venda_cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Início de variáveis padrão para relatório
Dim lLinha As Integer
Dim lPagina As Integer
'Fim de variáveis padrão para relatório
Dim l_quantidade As Currency
Dim l_valor As Currency
Dim l_sql As String

Private Cliente As New cCliente
Private BaixaNotaAbastecimento As New cBaixaNotaAbastecimento
Function ExisteNotaAbastecimento() As Boolean
    ExisteNotaAbastecimento = False
'    With tbl_baixa_nota_abastecimento
'        If .RecordCount > 0 Then
'            .Seek ">=", Cliente.Codigo, CDate(msk_data_i.Text), 0, 0, CDate("01/01/1900"), 1, "0"
'            If Not .NoMatch Then
'                If ![Codigo do Cliente] = Cliente.Codigo And ![Data do Abastecimento] <= CDate(msk_data_f) Then
                    ExisteNotaAbastecimento = True
'                End If
'            End If
'        End If
'    End With
End Function
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set BaixaNotaAbastecimento = Nothing
    Set Cliente = Nothing
End Sub
Private Sub Relatorio()
    Dim x_data_i As String
    Dim x_data_f As String
    Dim x_nome_empresa As String
    
    If ExisteNotaAbastecimento Then
        
        
        x_data_i = "date(" & Format(msk_data_i.Text, "yyyy,mm,dd") & ")"
        x_data_f = "date(" & Format(msk_data_f.Text, "yyyy,mm,dd") & ")"
        x_nome_empresa = g_nome_empresa
        If bdAccess Then
            cr_relato.ReportFileName = "\VB5\Sgp\Data\lst_venda_cliente.rpt"
            x_data_i = "date(" & Format(msk_data_i.Text, "yyyy,mm,dd") & ")"
            x_data_f = "date(" & Format(msk_data_f.Text, "yyyy,mm,dd") & ")"
            x_nome_empresa = """" & x_nome_empresa & """"
        ElseIf bdSqlServer Then
            cr_relato.ReportFileName = "\VB5\Sgp\Data\SQL_lst_venda_cliente.rpt"
            x_data_i = "date(" & Format(msk_data_i.Text, "yyyy,MM,dd") & ")"
            x_data_f = "date(" & Format(msk_data_f.Text, "yyyy,MM,dd") & ")"
            'x_nome_empresa = preparaTexto(x_nome_empresa)
        End If

'        cr_relato.SortFields(0) = "+{Produto.Nome}"
        cr_relato.Formulas(0) = "f_empresa = BeforeReadingRecords;""" & x_nome_empresa & """"
        cr_relato.Formulas(1) = "f_data_hoje =  BeforeReadingRecords;""" & msk_data & """"
        cr_relato.Formulas(2) = "f_data_inicial =  BeforeReadingRecords;""" & msk_data_i & """"
        cr_relato.Formulas(3) = "f_data_final =  BeforeReadingRecords;""" & msk_data_f & """"
        cr_relato.Formulas(4) = "f_cliente =  BeforeReadingRecords;""" & Cliente.RazaoSocial & """"
        l_sql = "{Baixa_Nota_Abastecimento.Codigo do Cliente} = " & Cliente.Codigo
        l_sql = l_sql & " And {Baixa_Nota_Abastecimento.Data do Abastecimento} >= " & x_data_i
        l_sql = l_sql & " And {Baixa_Nota_Abastecimento.Data do Abastecimento} <= " & x_data_f
        If bdSqlServer Then
            cr_relato.Connect = "DSN=sgp_data;UID=sa;PWD=" & gSenhaBD
            cr_relato.Password = gSenhaBD
        End If
        cr_relato.SelectionFormula = l_sql
        cr_relato.Action = 1
    Else
        MsgBox "O cliente " & Trim(Cliente.RazaoSocial) & "." & Chr(10) & "Não tem movimento no período " & msk_data_i & " à " & msk_data_f & ".", 64, "Mensagem para o operador!"
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        txt_cliente.SetFocus
    Else
        msk_data = RetiraGString(1)
        txt_cliente.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_data_f_Click()
    g_string = msk_data_f
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
    Else
        msk_data_f = RetiraGString(1)
    End If
    g_string = " "
    cmd_visualizar.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cmd_visualizar.SetFocus
    Else
        msk_data_i = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    If ValidaCampos Then
        If SelecionaImpressoraHP(Me) Then
            cr_relato.Destination = 1
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        End If
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data de emissão.", 64, "Atenção!"
        msk_data.SetFocus
    ElseIf Not IsDate(msk_data_i) Then
        MsgBox "Informe a data inicial.", 64, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f) Then
        MsgBox "Informe a data final.", 64, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f) < CDate(msk_data_i) Then
        MsgBox "Data final deve ser maior ou igual a " & CDate(msk_data_i) & ".", 64, "Atenção!"
        msk_data_f.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub ZZ_NAO_SEI()
'    If cbo_funcionario.ListIndex = -1 Or chk_acumulado.Value = 1 Then
'        If tbl_tabela_premiacao.RecordCount > 0 Then
'            tbl_tabela_premiacao.Seek "=", CDate("01/" & Format(msk_data_i, "mm") & "/" & Format(msk_data_i, "yyyy"))
'            If tbl_tabela_premiacao.NoMatch Then
'                If (MsgBox("Não existe o registro " & Format(msk_data_i, "mm") & "/" & Format(msk_data_i, "yyyy") & " cadastrado na tabela de premiação." & Chr(10) & "A premiação não será calculada." & Chr(10) & "Imprime com a premiação zerada?", 4 + 32 + 256, "Erro de Verificação!")) = 7 Then
'                    cmd_sair.SetFocus
'                    Exit Sub
'                End If
'            End If
'        End If
'    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    If ValidaCampos Then
        If SelecionaImpressoraHP(Me) Then
            cr_relato.Destination = 0
            Call GravaAuditoria(1, Me.name, 6, "")
            Relatorio
        End If
    End If
End Sub
Private Sub dtcboCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_i.SetFocus
    End If
End Sub
Private Sub dtcboCliente_LostFocus()
    If dtcboCliente.BoundText <> "" Then
        If Cliente.LocalizarCodigo(Val(dtcboCliente.BoundText)) Then
            If Cliente.Codigo <> Val(txt_cliente.Text) Then
                txt_cliente.Text = Cliente.Codigo
            End If
            msk_data_i.SetFocus
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        msk_data_i.Text = Format(CDate(g_data_def), "dd/mm/yyyy")
        msk_data_f.Text = Format(CDate(g_data_def), "dd/mm/yyyy")
        txt_cliente.SetFocus
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_imprimir_Click
    ElseIf KeyCode = vbKeyF9 Then
        KeyCode = 0
        cmd_visualizar_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    Set adodcCliente.Recordset = Conectar.RsConexao("SELECT Codigo, [Razao Social] FROM Cliente WHERE Inativo = " & preparaBooleano(False) & " ORDER BY [Razao Social]")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_f_GotFocus()
    msk_data_f.SelStart = 0
    msk_data_f.SelLength = 2
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub msk_data_i_GotFocus()
    msk_data_i.SelStart = 0
    msk_data_i.SelLength = 2
End Sub
Private Sub msk_data_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_f.SetFocus
    End If
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_cliente.SetFocus
    End If
End Sub
Private Sub txt_cliente_GotFocus()
    txt_cliente.SelStart = 0
    txt_cliente.SelLength = Len(txt_cliente)
End Sub
Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboCliente.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_cliente_LostFocus()
    If Val(txt_cliente) > 0 Then
        If Cliente.LocalizarCodigo(Val(txt_cliente.Text)) Then
            dtcboCliente.BoundText = Val(txt_cliente.Text)
            msk_data_i.SetFocus
            Exit Sub
        Else
            MsgBox "Cliente não cadastro.", vbInformation, "Atenção!"
            txt_cliente.SetFocus
            Exit Sub
        End If
    End If
End Sub
