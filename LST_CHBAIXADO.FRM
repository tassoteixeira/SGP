VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form lst_cheque_baixados 
   Caption         =   "Relação dos Cheques Baixados"
   ClientHeight    =   3495
   ClientLeft      =   1395
   ClientTop       =   2190
   ClientWidth     =   6975
   Icon            =   "LST_CHBAIXADO.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "LST_CHBAIXADO.frx":030A
   ScaleHeight     =   3495
   ScaleWidth      =   6975
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4920
      Picture         =   "LST_CHBAIXADO.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2520
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3060
      Picture         =   "LST_CHBAIXADO.frx":19E2
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Imprime cheques baixados."
      Top             =   2520
      Width           =   795
   End
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1200
      Picture         =   "LST_CHBAIXADO.frx":2FEC
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Visualiza cheques baixados."
      Top             =   2520
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6735
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2700
         Picture         =   "LST_CHBAIXADO.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2700
         Picture         =   "LST_CHBAIXADO.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "LST_CHBAIXADO.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbo_tipo_movimento 
         Height          =   315
         Left            =   3420
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1260
         Width           =   3135
      End
      Begin VB.TextBox txt_periodo_f 
         Height          =   285
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   13
         Top             =   1920
         Width           =   195
      End
      Begin VB.TextBox txt_periodo_i 
         Height          =   285
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   11
         Top             =   1500
         Width           =   195
      End
      Begin VB.Frame frmTipoData 
         Caption         =   "Tipo de Data"
         Height          =   915
         Left            =   3420
         TabIndex        =   14
         Top             =   180
         Width           =   1455
         Begin VB.OptionButton opt_data 
            Caption         =   "Vencimento"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   540
            Width           =   1215
         End
         Begin VB.OptionButton opt_data 
            Caption         =   "Emissão"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
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
      Begin VB.Label Label4 
         Caption         =   "Período Final"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1860
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Período Inicial"
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
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\VB5\Sgp\Data\lst_chbaixado.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
End
Attribute VB_Name = "lst_cheque_baixados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_sql As String
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
End Sub
Private Sub PreencheCboTipoMovimento()
    cbo_tipo_movimento.Clear
    cbo_tipo_movimento.AddItem "0 - ** Todos os Caixas"
    cbo_tipo_movimento.AddItem "1 - Caixa de Combustíveis"
    cbo_tipo_movimento.AddItem "2 - Caixa de Óleos/Diversos"
End Sub
Private Sub Relatorio()
    Dim x_data As String
    Dim x_data_i As String
    Dim x_data_f As String
    Dim x_tipo_movimento As String
    
    If bdAccess Then
        cr_relato.ReportFileName = "\VB5\Sgp\Data\lst_chbaixado.rpt"
        x_data_i = "date(" & Format(msk_data_i.Text, "yyyy,mm,dd") & ")"
        x_data_f = "date(" & Format(msk_data_f.Text, "yyyy,mm,dd") & ")"
        x_tipo_movimento = Mid(cbo_tipo_movimento.Text, 1, 1)
    ElseIf bdSqlServer Then
        cr_relato.ReportFileName = "\VB5\Sgp\Data\SQL_lst_chbaixado.rpt"
        x_data_i = "date(" & Format(msk_data_i.Text, "yyyy,MM,dd") & ")"
        x_data_f = "date(" & Format(msk_data_f.Text, "yyyy,MM,dd") & ")"
        x_tipo_movimento = preparaTexto(Mid(cbo_tipo_movimento.Text, 1, 1))
    End If
    If opt_data(0) Then
        x_data = "{Baixa_Cheque.Data de Emissao}"
    Else
        x_data = "{Baixa_Cheque.Data do Vencimento}"
    End If
    If x_tipo_movimento = "" Then
        x_tipo_movimento = "0"
    End If
    cr_relato.Formulas(0) = "f_empresa = BeforeReadingRecords;""" & g_nome_empresa & """"
    cr_relato.Formulas(1) = "f_data_hoje =  BeforeReadingRecords;""" & msk_data & """"
    cr_relato.Formulas(2) = "f_data_inicial =  BeforeReadingRecords;""" & msk_data_i & """"
    cr_relato.Formulas(3) = "f_data_final =  BeforeReadingRecords;""" & msk_data_f & """"
    cr_relato.Formulas(4) = "f_tipo_movimento =  BeforeReadingRecords;""" & Mid(cbo_tipo_movimento.Text, 5, Len(cbo_tipo_movimento.Text) - 4) & """"
    cr_relato.Formulas(5) = "f_periodo_inicial =  BeforeReadingRecords;""" & txt_periodo_i & """"
    cr_relato.Formulas(6) = "f_periodo_final =  BeforeReadingRecords;""" & txt_periodo_f & """"
    l_sql = "{Baixa_Cheque.Empresa} = " & g_empresa
    l_sql = l_sql & " And " & x_data & " >= " & x_data_i
    l_sql = l_sql & " And " & x_data & " <= " & x_data_f
    l_sql = l_sql & " And {Baixa_Cheque.Periodo} >= " & preparaTexto(txt_periodo_i.Text)
    l_sql = l_sql & " And {Baixa_Cheque.Periodo} <= " & preparaTexto(txt_periodo_f.Text)
    l_sql = l_sql & " And ({Baixa_Cheque.Tipo do Movimento} = " & x_tipo_movimento
    l_sql = l_sql & " Or " & x_tipo_movimento & " = " & preparaTexto("0") & ")"
    If bdSqlServer Then
        cr_relato.Connect = "DSN=sgp_data;UID=sa;PWD=" & gSenhaBD
        cr_relato.Password = gSenhaBD
    End If
    cr_relato.SelectionFormula = l_sql
    cr_relato.Action = 1
End Sub
Private Sub cbo_tipo_movimento_GotFocus()
    SendMessageLong cbo_tipo_movimento.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_tipo_movimento_KeyPress(KeyAscii As Integer)
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
        txt_periodo_i.SetFocus
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
    txt_periodo_i.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        txt_periodo_i.SetFocus
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
    Dim x_data_i As Date
    Dim x_data_f As Date
    x_data_i = msk_data_i
    x_data_f = msk_data_f
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data de emissão.", 64, "Atenção!"
        msk_data.SetFocus
    ElseIf Not IsDate(msk_data_i) Then
        MsgBox "Informe a data inicial.", 64, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f) Then
        MsgBox "Informe a data final.", 64, "Atenção!"
        msk_data_f.SetFocus
    ElseIf x_data_f < x_data_i Then
        MsgBox "Data final deve ser maior que a data inicial.", 64, "Atenção!"
        msk_data_f.SetFocus
    ElseIf Val(txt_periodo_i) < 1 Or Val(txt_periodo_i) > 4 Then
        MsgBox "Informe o período inicial de '1' a '4'.", 64, "Atenção!"
        txt_periodo_i.SetFocus
    ElseIf Val(txt_periodo_f) < 1 Or Val(txt_periodo_f) > 4 Then
        MsgBox "Informe o período final de '1' a '4'.", 64, "Atenção!"
        txt_periodo_f.SetFocus
    ElseIf Val(txt_periodo_f) < Val(txt_periodo_i) Then
        MsgBox "O período final deve ser maior.", 64, "Atenção!"
        txt_periodo_f.SetFocus
    ElseIf cbo_tipo_movimento.ListIndex = -1 Then
        MsgBox "Selecione tipo de movimento na caixa combo.", 64, "Atenção!"
        cbo_tipo_movimento.SetFocus
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
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Date
        PreencheCboTipoMovimento
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
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_periodo_i.SetFocus
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
Private Sub opt_data_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo_movimento.SetFocus
    End If
End Sub
Private Sub txt_periodo_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        opt_data(0).SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_periodo_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_periodo_f.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
