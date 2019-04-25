VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_funcionario 
   Caption         =   "Emissão de Funcionário (Por Empresa)"
   ClientHeight    =   1875
   ClientLeft      =   1575
   ClientTop       =   3075
   ClientWidth     =   4650
   Icon            =   "lst_funcionario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1875
   ScaleWidth      =   4650
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   600
      Picture         =   "lst_funcionario.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Visualiza relação de funcionários."
      Top             =   900
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   1920
      Picture         =   "lst_funcionario.frx":1A24
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprime relação de funcionários."
      Top             =   900
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3240
      Picture         =   "lst_funcionario.frx":302E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   900
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4395
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_funcionario.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
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
      Left            =   120
      Top             =   900
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\VB5\Sgp\Data\lst_funcionario.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_funcionario"
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
        cr_relato.SortFields(0) = "+{Funcionario.Nome}"
        cr_relato.Formulas(0) = "f_empresa = BeforeReadingRecords;""" & g_nome_empresa & """"
        cr_relato.Formulas(1) = "f_data_hoje =  BeforeReadingRecords;""" & msk_data & """"
        l_sql = "{Funcionario.Empresa} = " & g_empresa
        l_sql = l_sql & " And {Funcionario.Situacao} = 'A'"
        cr_relato.SelectionFormula = l_sql
        cr_relato.Action = 1
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = RetiraGString(1)
    cmd_visualizar.SetFocus
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
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        msk_data.SetFocus
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
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
