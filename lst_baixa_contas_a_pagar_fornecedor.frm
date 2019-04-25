VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_baixa_contas_a_pagar_fornecedor 
   Caption         =   "Relação de Baixas das Contas à Pagar por Fornecedor"
   ClientHeight    =   3135
   ClientLeft      =   1950
   ClientTop       =   1890
   ClientWidth     =   6795
   Icon            =   "lst_baixa_contas_a_pagar_fornecedor.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_baixa_contas_a_pagar_fornecedor.frx":030A
   ScaleHeight     =   3135
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1200
      Picture         =   "lst_baixa_contas_a_pagar_fornecedor.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Visualiza baixa de contas à pagar por fornecedor."
      Top             =   2160
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_baixa_contas_a_pagar_fornecedor.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Imprime baixa de contas à pagar por fornecedor."
      Top             =   2160
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4800
      Picture         =   "lst_baixa_contas_a_pagar_fornecedor.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2160
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_baixa_contas_a_pagar_fornecedor.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_baixa_contas_a_pagar_fornecedor.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_baixa_contas_a_pagar_fornecedor.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin VB.ComboBox cbo_fornecedor 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1500
         Width           =   4755
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
         Caption         =   "Fornecedor"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1500
         Width           =   1515
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
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1515
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
      Left            =   120
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\VB5\Sgp\Data\lst_baixa_contas_a_pagar_fornecedor.rpt"
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
Attribute VB_Name = "lst_baixa_contas_a_pagar_fornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lSQL As String
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
End Sub
Private Sub Relatorio()
    Dim x_data_i As String
    Dim x_data_f As String
    
    If bdAccess Then
        cr_relato.ReportFileName = "\VB5\Sgp\Data\lst_baixa_contas_a_pagar_fornecedor.rpt"
        x_data_i = "date(" & Format(msk_data_i.Text, "yyyy,mm,dd") & ")"
        x_data_f = "date(" & Format(msk_data_f.Text, "yyyy,mm,dd") & ")"
    ElseIf bdSqlServer Then
        cr_relato.ReportFileName = "\VB5\Sgp\Data\SQL_lst_baixa_contas_a_pagar_fornecedor.rpt"
        x_data_i = "date(" & Format(msk_data_i.Text, "yyyy,MM,dd") & ")"
        x_data_f = "date(" & Format(msk_data_f.Text, "yyyy,MM,dd") & ")"
    End If
    cr_relato.Formulas(0) = "f_empresa = BeforeReadingRecords;""" & g_nome_empresa & """"
    cr_relato.Formulas(0) = "f_empresa = BeforeReadingRecords;""" & "Grupo X" & """"
    cr_relato.Formulas(1) = "f_data_hoje =  BeforeReadingRecords;""" & msk_data & """"
    cr_relato.Formulas(2) = "f_data_inicial =  BeforeReadingRecords;""" & msk_data_i & """"
    cr_relato.Formulas(3) = "f_data_final =  BeforeReadingRecords;""" & msk_data_f & """"
    cr_relato.Formulas(4) = "f_fornecedor =  BeforeReadingRecords;""" & cbo_fornecedor.Text & """"
    lSQL = "{Baixa_Pagar.data_pagamento} >= " & x_data_i
    lSQL = lSQL & " And {Baixa_Pagar.data_pagamento} <= " & x_data_f
    If cbo_fornecedor.ItemData(cbo_fornecedor.ListIndex) > 0 Then
        lSQL = lSQL & " And {Baixa_Pagar.codigo_fornecedor} = " & cbo_fornecedor.ItemData(cbo_fornecedor.ListIndex)
    End If
    If bdSqlServer Then
        cr_relato.Connect = "DSN=sgp_data;UID=sa;PWD=" & gSenhaBD
        cr_relato.Password = gSenhaBD
    End If
    cr_relato.SelectionFormula = lSQL
    cr_relato.Action = 1
End Sub
Private Sub cbo_fornecedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cbo_fornecedor.SetFocus
    Else
        msk_data.Text = RetiraGString(1)
        msk_data_i.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_data_f_Click()
    g_string = msk_data_f.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
    Else
        msk_data_f.Text = RetiraGString(1)
    End If
    g_string = ""
    cbo_fornecedor.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cbo_fornecedor.SetFocus
    Else
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = ""
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
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf Not IsDate(msk_data_i.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f.Text) < CDate(msk_data_i.Text) Then
        MsgBox "Data final deve ser maior ou igual a " & msk_data_i.Text & ".", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf cbo_fornecedor.ListIndex = -1 Then
        MsgBox "Selecione um fornecedor.", vbInformation, "Atenção!"
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
        msk_data_i.Text = Format(g_data_def - 30, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def, "dd/mm/yyyy")
        cbo_fornecedor.ListIndex = 0
        msk_data_i.SetFocus
    End If
    Screen.MousePointer = 1
End Sub
Private Sub PreencheCboFornecedor()
    Dim rstFornecedor As New adodb.Recordset
    
    cbo_fornecedor.Clear
    cbo_fornecedor.AddItem "Todos os Fornecedores"
    cbo_fornecedor.ItemData(cbo_fornecedor.NewIndex) = 0
    Set rstFornecedor = Conectar.RsConexao("SELECT Codigo, Nome FROM Fornecedor ORDER BY NOME")
    Do Until rstFornecedor.EOF
        cbo_fornecedor.AddItem rstFornecedor!Nome
        cbo_fornecedor.ItemData(cbo_fornecedor.NewIndex) = rstFornecedor!Codigo
        rstFornecedor.MoveNext
    Loop
    rstFornecedor.Close
    Set rstFornecedor = Nothing
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
        cbo_fornecedor.SetFocus
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
