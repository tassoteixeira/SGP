VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_contas_pagar 
   Caption         =   "Relação de Contas à Pagar"
   ClientHeight    =   3555
   ClientLeft      =   1950
   ClientTop       =   1890
   ClientWidth     =   6795
   Icon            =   "LST_CONTPAG.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "LST_CONTPAG.frx":030A
   ScaleHeight     =   3555
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "LST_CONTPAG.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Visualiza contas à pagar."
      Top             =   2580
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "LST_CONTPAG.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Imprime contas à pagar."
      Top             =   2580
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "LST_CONTPAG.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2580
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "LST_CONTPAG.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "LST_CONTPAG.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2760
         Picture         =   "LST_CONTPAG.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin VB.ComboBox cbo_fornecedor 
         Height          =   300
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1920
         Width           =   4755
      End
      Begin VB.ComboBox cbo_conta 
         Height          =   300
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1500
         Width           =   3435
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   1080
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data_i 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   660
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "F&ornecedor"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Conta"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Data &Inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Data &Final"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin Crystal.CrystalReport cr_relato 
      Left            =   180
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\VB5\Sgp\Data\lst_contas_pagar.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_contas_pagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbl_conta As Table
Dim tbl_fornecedor As Table
Dim l_sql As String
Private MovimentoContaPagar As New cMovimentoContaPagar
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_conta.Close
    tbl_fornecedor.Close
    Set MovimentoContaPagar = Nothing
End Sub
Private Sub PreencheCboConta()
    cbo_conta.Clear
    cbo_conta.AddItem "Todas as Contas"
    cbo_conta.ItemData(cbo_conta.NewIndex) = 0
    tbl_conta.Index = "id_nome"
    If tbl_conta.RecordCount > 0 Then
        tbl_conta.MoveFirst
        tbl_conta.Seek ">", g_empresa, " "
        Do Until tbl_conta.EOF
        
            If tbl_conta!Empresa <> g_empresa Then
                Exit Do
            End If

            cbo_conta.AddItem tbl_conta!Nome
            cbo_conta.ItemData(cbo_conta.NewIndex) = tbl_conta!Codigo
            tbl_conta.MoveNext
        Loop
    End If
    tbl_conta.Index = "id_codigo"
End Sub
Private Sub Relatorio()
    Dim x_data_i As String
    Dim x_data_f As String
    x_data_i = "date(" & Format(msk_data_i, "yyyy,mm,dd") & ")"
    x_data_f = "date(" & Format(msk_data_f, "yyyy,mm,dd") & ")"
    cr_relato.SortFields(0) = "+{contas_pagar.data_vencimento}"
    cr_relato.SortFields(1) = "+{contas_pagar.nome_fornecedor}"
    cr_relato.SortFields(2) = "+{contas_pagar.registro}"
    cr_relato.Formulas(0) = "f_empresa = BeforeReadingRecords;""" & g_nome_empresa & """"
    cr_relato.Formulas(1) = "f_data_hoje =  BeforeReadingRecords;""" & msk_data & """"
    cr_relato.Formulas(2) = "f_data_inicial =  BeforeReadingRecords;""" & msk_data_i & """"
    cr_relato.Formulas(3) = "f_data_final =  BeforeReadingRecords;""" & msk_data_f & """"
    cr_relato.Formulas(4) = "f_conta =  BeforeReadingRecords;""" & cbo_conta.Text & """"
    cr_relato.Formulas(5) = "f_fornecedor =  BeforeReadingRecords;""" & cbo_fornecedor.Text & """"
    l_sql = "{contas_pagar.empresa} = " & g_empresa
    l_sql = l_sql & " And {contas_pagar.data_vencimento} >= " & x_data_i
    l_sql = l_sql & " And {contas_pagar.data_vencimento} <= " & x_data_f
    If cbo_conta.ItemData(cbo_conta.ListIndex) > 0 Then
        l_sql = l_sql & " And {contas_pagar.codigo_conta} = " & cbo_conta.ItemData(cbo_conta.ListIndex)
    End If
    If cbo_fornecedor.ItemData(cbo_fornecedor.ListIndex) > 0 Then
        l_sql = l_sql & " And {contas_pagar.codigo_fornecedor} = " & cbo_fornecedor.ItemData(cbo_fornecedor.ListIndex)
    End If
    cr_relato.SelectionFormula = l_sql
    cr_relato.Action = 1
End Sub
Private Sub cbo_conta_GotFocus()
    SendMessageLong cbo_conta.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_conta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_fornecedor.SetFocus
    End If
End Sub
Private Sub cbo_fornecedor_GotFocus()
    SendMessageLong cbo_fornecedor.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_fornecedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cbo_conta.SetFocus
    Else
        msk_data = RetiraGString(1)
        msk_data_i.SetFocus
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
    cbo_conta.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cbo_conta.SetFocus
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
            cmd_sair.SetFocus
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
        MsgBox "Data final deve ser maior ou igual a " & msk_data_i & ".", 64, "Atenção!"
        msk_data_f.SetFocus
    ElseIf cbo_conta.ListIndex = -1 Then
        MsgBox "Selecione uma conta.", 64, "Atenção!"
        cbo_conta.SetFocus
    ElseIf cbo_fornecedor.ListIndex = -1 Then
        MsgBox "Selecione um fornecedor.", 64, "Atenção!"
        cbo_fornecedor.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
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
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i.Text = Format(MovimentoContaPagar.PrimeiraDataVencimento(g_empresa), "dd/mm/yyyy")  '  Format(g_data_def, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def + 30, "dd/mm/yyyy")
        cbo_conta.ListIndex = 0
        cbo_fornecedor.ListIndex = 0
    End If
    msk_data_i.SetFocus
End Sub
Private Sub PreencheCboFornecedor()
    cbo_fornecedor.Clear
    cbo_fornecedor.AddItem "Todos os Fornecedores"
    cbo_fornecedor.ItemData(cbo_fornecedor.NewIndex) = 0
    tbl_fornecedor.Index = "id_nome"
    If tbl_fornecedor.RecordCount > 0 Then
        tbl_fornecedor.Seek ">", g_empresa, " "
        Do Until tbl_fornecedor.EOF
            If tbl_fornecedor!Empresa <> g_empresa Then
                Exit Do
            End If
            cbo_fornecedor.AddItem tbl_fornecedor!Nome
            cbo_fornecedor.ItemData(cbo_fornecedor.NewIndex) = tbl_fornecedor!Codigo
            tbl_fornecedor.MoveNext
        Loop
    End If
    tbl_fornecedor.Index = "id_codigo"
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
    Screen.MousePointer = 1
    CentraForm Me
    Set tbl_conta = bd_sgp.OpenTable("contas")
    Set tbl_fornecedor = bd_sgp.OpenTable("Fornecedor")
    PreencheCboConta
    PreencheCboFornecedor
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_f_GotFocus()
    msk_data_f.SelStart = 0
    msk_data_f.SelLength = 5
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_conta.SetFocus
    End If
End Sub
Private Sub msk_data_i_GotFocus()
    msk_data_i.SelStart = 0
    msk_data_i.SelLength = 5
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
        msk_data_i.SetFocus
    End If
End Sub
