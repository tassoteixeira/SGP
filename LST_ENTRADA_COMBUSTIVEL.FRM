VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_entrada_combustivel 
   Caption         =   "Relação de Entrada de Combustível"
   ClientHeight    =   3195
   ClientLeft      =   2790
   ClientTop       =   3810
   ClientWidth     =   5475
   Icon            =   "LST_ENTRADA_COMBUSTIVEL.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "LST_ENTRADA_COMBUSTIVEL.frx":030A
   ScaleHeight     =   3195
   ScaleWidth      =   5475
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   840
      Picture         =   "LST_ENTRADA_COMBUSTIVEL.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Visualiza entrada de combustível."
      Top             =   2220
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2340
      Picture         =   "LST_ENTRADA_COMBUSTIVEL.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Imprime entrada de combustível."
      Top             =   2220
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3840
      Picture         =   "LST_ENTRADA_COMBUSTIVEL.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2220
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1995
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5235
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "LST_ENTRADA_COMBUSTIVEL.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2700
         Picture         =   "LST_ENTRADA_COMBUSTIVEL.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2700
         Picture         =   "LST_ENTRADA_COMBUSTIVEL.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin VB.ComboBox cbo_combustivel 
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
         Width           =   975
         _ExtentX        =   1720
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
         Width           =   975
         _ExtentX        =   1720
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
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "&Combustível"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin Crystal.CrystalReport cr_relato 
      Left            =   120
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\VB5\Sgp\Data\lst_entrada_combustivel.rpt"
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
Attribute VB_Name = "lst_entrada_combustivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_sql As String
Dim l_tipo_combustivel As String
Dim lCustoMedio As Currency

Private Combustivel As New cCombustivel
Private Sub AtivaBotoes(ByVal pAtiva As Boolean)
    cmd_visualizar.Enabled = pAtiva
    cmd_imprimir.Enabled = pAtiva
    cmd_sair.Enabled = pAtiva
    If pAtiva = False Then
        frmAguarde.Show
        Call frmAguarde.MostraMensagens("Gerando Relatório!", Me.Caption, Me.Top, Me.Left, Me.Width, Me.Height)
        DoEvents
    Else
        Call frmAguarde.Finaliza
    End If
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Combustivel = Nothing
End Sub
Private Sub Relatorio()
    Dim x_data_i As String
    Dim x_data_f As String
    Dim xTipoCombustivel As String
    Dim xNumeroNota As String
    
    If bdAccess Then
        cr_relato.ReportFileName = "\VB5\Sgp\Data\lst_entrada_combustivel_lmc.rpt"
        x_data_i = "date(" & Format(msk_data_i.Text, "yyyy,mm,dd") & ")"
        x_data_f = "date(" & Format(msk_data_f.Text, "yyyy,mm,dd") & ")"
        xTipoCombustivel = """" & l_tipo_combustivel & """"
        xNumeroNota = """" & "1" & """"
    ElseIf bdSqlServer Then
        cr_relato.ReportFileName = "\VB5\Sgp\Data\SQL_lst_entrada_combustivel_lmc.rpt"
        x_data_i = "date(" & Format(msk_data_i.Text, "yyyy,MM,dd") & ")"
        x_data_f = "date(" & Format(msk_data_f.Text, "yyyy,MM,dd") & ")"
        xTipoCombustivel = preparaTexto(l_tipo_combustivel)
        xNumeroNota = preparaTexto("1")
    End If
    If g_nome_usuario = "L.M.C." Then
        cr_relato.SortFields(0) = "+{Entrada_Combustivel_LMC.Data}"
        cr_relato.SortFields(1) = "+{Entrada_Combustivel_LMC.Numero da Nota}"
        cr_relato.Formulas(0) = "f_empresa = BeforeReadingRecords;""" & g_nome_empresa & """"
        cr_relato.Formulas(1) = "f_data_hoje =  BeforeReadingRecords;""" & msk_data & """"
        cr_relato.Formulas(2) = "f_data_inicial =  BeforeReadingRecords;""" & msk_data_i & """"
        cr_relato.Formulas(3) = "f_data_final =  BeforeReadingRecords;""" & msk_data_f & """"
        cr_relato.Formulas(4) = "f_combustivel =  BeforeReadingRecords;""" & cbo_combustivel.Text & """"
        l_sql = "{Entrada_Combustivel_LMC.Empresa} = " & g_empresa
        l_sql = l_sql & " And {Entrada_Combustivel_LMC.Data} >= " & x_data_i
        l_sql = l_sql & " And {Entrada_Combustivel_LMC.Data} <= " & x_data_f
        'If cbo_combustivel.ItemData(cbo_combustivel.ListIndex) > 0 Then
        l_sql = l_sql & " And {Entrada_Combustivel_LMC.Tipo de Combustivel} = " & xTipoCombustivel
        l_sql = l_sql & " And {Entrada_Combustivel_LMC.Numero da Nota} <> " & xNumeroNota
        'End If
    Else
        cr_relato.SortFields(0) = "+{Entrada_Combustivel.Data}"
        cr_relato.SortFields(1) = "+{Entrada_Combustivel.Numero da Nota}"
        cr_relato.Formulas(0) = "f_empresa = BeforeReadingRecords;""" & g_nome_empresa & """"
        cr_relato.Formulas(1) = "f_data_hoje =  BeforeReadingRecords;""" & msk_data & """"
        cr_relato.Formulas(2) = "f_data_inicial =  BeforeReadingRecords;""" & msk_data_i & """"
        cr_relato.Formulas(3) = "f_data_final =  BeforeReadingRecords;""" & msk_data_f & """"
        cr_relato.Formulas(4) = "f_combustivel =  BeforeReadingRecords;""" & cbo_combustivel.Text & """"
        l_sql = "{Entrada_Combustivel.Empresa} = " & g_empresa
        l_sql = l_sql & " And {Entrada_Combustivel.Data} >= " & x_data_i
        l_sql = l_sql & " And {Entrada_Combustivel.Data} <= " & x_data_f
        'If cbo_combustivel.ItemData(cbo_combustivel.ListIndex) > 0 Then
        l_sql = l_sql & " And {Entrada_Combustivel.Tipo de Combustivel} = " & xTipoCombustivel
        l_sql = l_sql & " And {Entrada_Combustivel.Numero da Nota} <> " & xNumeroNota
        'End If
        If bdAccess Then
            cr_relato.ReportFileName = "\VB5\Sgp\Data\lst_entrada_combustivel.rpt"
        ElseIf bdSqlServer Then
            cr_relato.ReportFileName = "\VB5\Sgp\Data\SQL_lst_entrada_combustivel.rpt"
        End If
    End If
    If bdSqlServer Then
        cr_relato.Connect = "DSN=sgp_data;UID=sa;PWD=" & gSenhaBD
        cr_relato.Password = gSenhaBD
    End If
    cr_relato.SelectionFormula = l_sql
    cr_relato.Action = 1
End Sub
Private Sub cbo_combustivel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub cbo_combustivel_LostFocus()
    If cbo_combustivel.ListIndex <> -1 Then
        l_tipo_combustivel = Mid(cbo_combustivel, 1, 2)
        If Not Combustivel.LocalizarCodigo(g_empresa, l_tipo_combustivel) Then
            MsgBox "Combustível não cadastrado!", vbCritical, "Erro de Integridade!"
            cbo_combustivel.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cbo_combustivel.SetFocus
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
    cbo_combustivel.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cbo_combustivel.SetFocus
    Else
        msk_data_i = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    If ValidaCampos Then
        AtivaBotoes (False)
        If SelecionaImpressoraHP(Me) Then
            cr_relato.Destination = 1
            Call GravaAuditoria(1, Me.name, 7, "Ref:" & msk_data_i.Text & " a " & msk_data_f.Text & " Comb:" & cbo_combustivel.Text)
            Relatorio
        End If
        AtivaBotoes (True)
        cmd_sair.SetFocus
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
        MsgBox "Data final deve ser maior que a data inicial.", 64, "Atenção!"
        msk_data_f.SetFocus
    ElseIf cbo_combustivel.ListIndex = -1 Then
        MsgBox "Selecione o combustível.", 64, "Atenção!"
        cbo_combustivel.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    If ValidaCampos Then
        AtivaBotoes (False)
        If SelecionaImpressoraHP(Me) Then
            cr_relato.Destination = 0
            Call GravaAuditoria(1, Me.name, 6, "Ref:" & msk_data_i.Text & " a " & msk_data_f.Text & " Comb:" & cbo_combustivel.Text)
            Relatorio
        End If
        AtivaBotoes (True)
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        msk_data_i.Text = fDataPrimeiroDiaMesAnterior(Date)
        msk_data_f.Text = fDataUltimoDiaMesAnterior(Date)
        msk_data_i.SetFocus
    End If
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
    PreencheCboCombustivel
End Sub
Private Sub PreencheCboCombustivel()
    Dim rstCombustivel As New adodb.Recordset
    
    cbo_combustivel.Clear
    Set rstCombustivel = Conectar.RsConexao("SELECT Codigo, Nome FROM Combustivel WHERE Empresa = " & g_empresa & " ORDER BY Nome")
    'loop RecordSet
    With rstCombustivel
        If Not .BOF Or Not .EOF Then
            .MoveFirst
            Do Until .EOF
                cbo_combustivel.AddItem !Codigo & " - " & !Nome
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rstCombustivel = Nothing
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
        cbo_combustivel.SetFocus
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
