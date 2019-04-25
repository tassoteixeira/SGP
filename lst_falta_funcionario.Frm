VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form lst_falta_funcionario 
   Caption         =   "Relação das Faltas dos Funcionários"
   ClientHeight    =   3195
   ClientLeft      =   360
   ClientTop       =   4245
   ClientWidth     =   6795
   Icon            =   "lst_falta_funcionario.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_falta_funcionario.frx":030A
   ScaleHeight     =   3195
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_falta_funcionario.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Visualiza falta de funcionário."
      Top             =   2220
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_falta_funcionario.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Imprime falta de funcionário."
      Top             =   2220
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_falta_funcionario.frx":3074
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
      Width           =   6555
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_falta_funcionario.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_falta_funcionario.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_falta_funcionario.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin VB.Data dta_funcionario 
         Caption         =   "dta_funcionario"
         Connect         =   "Access"
         DatabaseName    =   "Sgp_data.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1500
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSDBCtls.DBCombo dbcbo_funcionario 
         Bindings        =   "lst_falta_funcionario.frx":7F94
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   1500
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "nome"
         BoundColumn     =   "codigo"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   1680
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
      Begin MSMask.MaskEdBox msk_data_i 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   660
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
         Left            =   1680
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
      Begin VB.Label Label3 
         Caption         =   "Funcionário"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Data Inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Data Final"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Data de Emissão"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1515
      End
   End
   Begin Crystal.CrystalReport cr_relato 
      Left            =   300
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\VB5\Sgp\Data\lst_falta_funcionario.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
End
Attribute VB_Name = "lst_falta_funcionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_sql As String
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
End Sub
Private Sub Relatorio()
    cr_relato.SortFields(0) = "+{Funcionario.Nome}"
    cr_relato.SortFields(1) = "+{Movimento_Falta_Funcionario.Data}"
    cr_relato.Formulas(0) = "f_empresa = BeforeReadingRecords;""" & g_nome_empresa & """"
    cr_relato.Formulas(1) = "f_data_hoje =  BeforeReadingRecords;""" & msk_data & """"
    cr_relato.Formulas(2) = "f_data_inicial =  BeforeReadingRecords;""" & msk_data_i & """"
    cr_relato.Formulas(3) = "f_data_final =  BeforeReadingRecords;""" & msk_data_f & """"
    l_sql = "{Movimento_Falta_Funcionario.Empresa} = " & g_empresa
    l_sql = l_sql & " And {Movimento_Falta_Funcionario.Data} >= " & "date(" & Format(msk_data_i, "yyyy,mm,dd") & ")"
    l_sql = l_sql & " And {Movimento_Falta_Funcionario.Data} <= " & "date(" & Format(msk_data_f, "yyyy,mm,dd") & ")"
    If Val(dbcbo_funcionario.BoundText) > 0 Then
        l_sql = l_sql & " And {Movimento_Falta_Funcionario.Codigo do Funcionario} = " & dbcbo_funcionario.BoundText
    End If
    cr_relato.SelectionFormula = l_sql
    cr_relato.Action = 1
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        dbcbo_funcionario.SetFocus
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
    dbcbo_funcionario.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        dbcbo_funcionario.SetFocus
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
        MsgBox "Data final deve ser maior ou igual que " & msk_data_i & ".", 64, "Atenção!"
        msk_data_f.SetFocus
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
Private Sub dbcbo_funcionario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        TabelaFuncionarioRefresh
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i.SetFocus
    End If
    Screen.MousePointer = 1
End Sub
Private Sub TabelaFuncionarioRefresh()
    dta_funcionario.RecordSource = "Select * From Funcionario Where Empresa = " & g_empresa & " And Situacao = 'A' Order By Nome"
    dta_funcionario.Refresh
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
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dbcbo_funcionario.SetFocus
    End If
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
