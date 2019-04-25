VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_duplicata_receber 
   Caption         =   "Relação de Duplicatas à Receber"
   ClientHeight    =   2895
   ClientLeft      =   3600
   ClientTop       =   3375
   ClientWidth     =   6795
   Icon            =   "lst_duplicata_receber.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_duplicata_receber.frx":030A
   ScaleHeight     =   2895
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_duplicata_receber.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Visualiza duplicatas à receber."
      Top             =   1920
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_duplicata_receber.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Imprime duplicatas à receber."
      Top             =   1920
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_duplicata_receber.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1920
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_duplicata_receber.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_duplicata_receber.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "lst_duplicata_receber.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.Frame frmOrdemRelatorio 
         Caption         =   "Ordem do Relatório"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   3855
         Begin VB.OptionButton opt_ordem 
            Caption         =   "Razão Social"
            Height          =   195
            Index           =   1
            Left            =   2160
            TabIndex        =   12
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton opt_ordem 
            Caption         =   "Data de Vencimento"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   1875
         End
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4800
         TabIndex        =   8
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
      Begin VB.Label Label1 
         Caption         =   "Data Inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Data Final"
         Height          =   315
         Left            =   3780
         TabIndex        =   7
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Data de Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin Crystal.CrystalReport cr_relato 
      Left            =   660
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\VB5\Sgp\Data\lst_duplicata_receber.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_duplicata_receber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_sql As String
Private MovDuplicataReceber As New cDuplicataReceber
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set MovDuplicataReceber = Nothing
End Sub
Private Sub Relatorio()
    Dim x_data_i As String
    Dim x_data_f As String
    
    If bdAccess Then
        cr_relato.ReportFileName = "\VB5\Sgp\Data\lst_duplicata_receber.rpt"
        x_data_i = "date(" & Format(msk_data_i.Text, "yyyy,mm,dd") & ")"
        x_data_f = "date(" & Format(msk_data_f.Text, "yyyy,mm,dd") & ")"
    ElseIf bdSqlServer Then
        cr_relato.ReportFileName = "\VB5\Sgp\Data\SQL_lst_duplicata_receber.rpt"
        x_data_i = "date(" & Format(msk_data_i.Text, "yyyy,MM,dd") & ")"
        x_data_f = "date(" & Format(msk_data_f.Text, "yyyy,MM,dd") & ")"
    End If
    If opt_ordem(0) Then
        cr_relato.SortFields(0) = "+{Duplicata_Receber.Data do Vencimento}"
    Else
        cr_relato.SortFields(0) = "+{Cliente.Razao Social}"
    End If
    cr_relato.Formulas(0) = "f_empresa = BeforeReadingRecords;""" & g_nome_empresa & """"
    cr_relato.Formulas(1) = "f_data_hoje =  BeforeReadingRecords;""" & msk_data & """"
    cr_relato.Formulas(2) = "f_data_inicial =  BeforeReadingRecords;""" & msk_data_i & """"
    cr_relato.Formulas(3) = "f_data_final =  BeforeReadingRecords;""" & msk_data_f & """"
    l_sql = "{Duplicata_Receber.Empresa} = " & g_empresa
    l_sql = l_sql & " And {Duplicata_Receber.Data do Vencimento} >= " & x_data_i
    l_sql = l_sql & " And {Duplicata_Receber.Data do Vencimento} <= " & x_data_f
    
    If bdSqlServer Then
        cr_relato.Connect = "DSN=sgp_data;UID=sa;PWD=" & gSenhaBD
        cr_relato.Password = gSenhaBD
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
        opt_ordem(0).SetFocus
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
    opt_ordem(0).SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        opt_ordem(0).SetFocus
    Else
        msk_data_i = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    cr_relato.Destination = 1
    If ValidaCampos Then
        If SelecionaImpressoraHP(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "Ref:" & msk_data_i.Text & " a " & msk_data_f.Text)
            Relatorio
        End If
    End If
    cmd_sair.SetFocus
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
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    cr_relato.Destination = 0
    If ValidaCampos Then
        If SelecionaImpressoraHP(Me) Then
            Call GravaAuditoria(1, Me.name, 6, "Ref:" & msk_data_i.Text & " a " & msk_data_f.Text)
            Relatorio
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i.Text = Format(MovDuplicataReceber.LocalizaPrimeiroVencimento(g_empresa), "dd/mm/yyyy")
        msk_data_f.Text = Format(MovDuplicataReceber.LocalizaUltimoVencimento(g_empresa), "dd/mm/yyyy")
        msk_data_i.SetFocus
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
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        opt_ordem(0).SetFocus
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
Private Sub opt_ordem_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
