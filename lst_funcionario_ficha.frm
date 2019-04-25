VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form emissao_funcionario_ficha 
   Caption         =   "Emissão de Funcionário (Ficha Individual)"
   ClientHeight    =   2295
   ClientLeft      =   1605
   ClientTop       =   3540
   ClientWidth     =   7530
   Icon            =   "lst_funcionario_ficha.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2295
   ScaleWidth      =   7530
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1320
      Picture         =   "lst_funcionario_ficha.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Visualiza ficha individual de funcionário."
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3360
      Picture         =   "lst_funcionario_ficha.frx":1A24
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprime ficha individual de funcionário."
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5400
      Picture         =   "lst_funcionario_ficha.frx":302E
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1320
      Width           =   795
   End
   Begin VB.Data dta_funcionario 
      Caption         =   "dta_funcionario"
      Connect         =   "Access"
      DatabaseName    =   "Sgp_data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4380
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Funcionario"
      Top             =   780
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.Frame frm_dados 
      Height          =   1155
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7275
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_funcionario_ficha.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optAdvertencia 
         Caption         =   "Sim"
         Height          =   195
         Index           =   0
         Left            =   5760
         TabIndex        =   5
         Top             =   300
         Visible         =   0   'False
         Width           =   555
      End
      Begin MSDBCtls.DBCombo dbcbo_funcionario 
         Bindings        =   "lst_funcionario_ficha.frx":599A
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   720
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   ""
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
      Begin VB.OptionButton optAdvertencia 
         Caption         =   "Não"
         Height          =   195
         Index           =   1
         Left            =   6420
         TabIndex        =   6
         Top             =   300
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbl_advertencia 
         Caption         =   "Emite Advertência/Suspenção"
         Height          =   255
         Left            =   3420
         TabIndex        =   4
         Top             =   300
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "&Funcionário"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   7
         Top             =   720
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
      Left            =   180
      Top             =   1260
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\VB5\Sgp\Data\lst_funcionario_ficha.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   180
      Top             =   1740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_funcionario_ficha"
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
    If ValidaCampos Then
        cr_relato.ReportFileName = "lst_funcionario_ficha.rpt"
        If optAdvertencia(0).Visible Then
            If optAdvertencia(0) Then
                cr_relato.ReportFileName = "lst_funcionario_ficha2.rpt"
            End If
        End If
        cr_relato.Formulas(0) = "f_empresa = BeforeReadingRecords;""" & g_nome_empresa & """"
        cr_relato.Formulas(1) = "f_data_hoje =  BeforeReadingRecords;""" & msk_data & """"
        l_sql = "{Funcionario.Empresa} = " & g_empresa
        l_sql = l_sql & " and {Funcionario.Codigo} = " & Val(dbcbo_funcionario.BoundText)
        cr_relato.SelectionFormula = l_sql
        cr_relato.Action = 1
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = RetiraGString(1)
    dbcbo_funcionario.SetFocus
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
    ElseIf Val(dbcbo_funcionario.BoundText) = 0 Then
        MsgBox "Selecione um funcionario.", 64, "Atenção!"
        dbcbo_funcionario.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Finaliza
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
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    TabelaFuncionarioRefresh
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        dbcbo_funcionario.SetFocus
    End If
    If optAdvertencia(0).Visible Then
        optAdvertencia(0) = True
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
    If g_nivel_acesso <= 3 Then
        lbl_advertencia.Visible = True
        optAdvertencia(1).Visible = True
        optAdvertencia(0).Visible = True
    End If
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        optAdvertencia(0).SetFocus
    End If
End Sub
Private Sub optAdvertencia_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dbcbo_funcionario.SetFocus
    End If
End Sub
