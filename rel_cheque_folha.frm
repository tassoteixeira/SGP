VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form relatorio_cheque_folha 
   Caption         =   "Relação dos Cheques (Avulso)"
   ClientHeight    =   4020
   ClientLeft      =   1875
   ClientTop       =   1725
   ClientWidth     =   6435
   Icon            =   "rel_cheque_folha.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "rel_cheque_folha.frx":030A
   ScaleHeight     =   4020
   ScaleWidth      =   6435
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4560
      Picture         =   "rel_cheque_folha.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3060
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2820
      Picture         =   "rel_cheque_folha.frx":19E2
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Imprime cheques (avulso)."
      Top             =   3060
      Width           =   795
   End
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1080
      Picture         =   "rel_cheque_folha.frx":2FEC
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Visualiza cheques (avulso)."
      Top             =   3060
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2895
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6315
      Begin VB.ComboBox cboChequePosse 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2460
         Width           =   2175
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "rel_cheque_folha.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "rel_cheque_folha.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2760
         Picture         =   "rel_cheque_folha.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin VB.Frame frmOrdem 
         Height          =   435
         Left            =   1680
         TabIndex        =   17
         Top             =   1920
         Width           =   4455
         Begin VB.OptionButton optOrdem 
            Caption         =   "Número do Cheque"
            Height          =   195
            Index           =   1
            Left            =   2640
            TabIndex        =   19
            Top             =   150
            Width           =   1755
         End
         Begin VB.OptionButton optOrdem 
            Caption         =   "Data de Emissão"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   150
            Width           =   1575
         End
      End
      Begin VB.Frame frmSituacao 
         Height          =   435
         Left            =   1680
         TabIndex        =   11
         Top             =   1440
         Width           =   4455
         Begin VB.OptionButton optSituacao 
            Caption         =   "Aberto"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   150
            Width           =   1035
         End
         Begin VB.OptionButton optSituacao 
            Caption         =   "Baixado"
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   13
            Top             =   150
            Width           =   1035
         End
         Begin VB.OptionButton optSituacao 
            Caption         =   "Cancelado"
            Height          =   195
            Index           =   2
            Left            =   2250
            TabIndex        =   14
            Top             =   150
            Width           =   1155
         End
         Begin VB.OptionButton optSituacao 
            Caption         =   "Todos"
            Height          =   195
            Index           =   3
            Left            =   3510
            TabIndex        =   15
            Top             =   150
            Width           =   915
         End
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
      Begin VB.Label Label11 
         Caption         =   "Cheque em posse"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2460
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Ordenar por"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "S&ituação"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Data Inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Data Final"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Data de Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport cr_relato 
      Left            =   60
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\VB5\Sgp\Data\rel_cheque_folha.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
End
Attribute VB_Name = "relatorio_cheque_folha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_sql As String
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_cheque.Close
End Sub
Private Sub PreencheCboChequePosse()
    cboChequePosse.Clear
    cboChequePosse.AddItem "0 Geral"
    cboChequePosse.ItemData(cboChequePosse.NewIndex) = 0
    cboChequePosse.AddItem "1 Favorecido"
    cboChequePosse.ItemData(cboChequePosse.NewIndex) = 1
    cboChequePosse.AddItem "2 Escritorio"
    cboChequePosse.ItemData(cboChequePosse.NewIndex) = 2
    cboChequePosse.AddItem "3 Pista"
    cboChequePosse.ItemData(cboChequePosse.NewIndex) = 3
End Sub
Private Sub Relatorio()
    Dim x_data As String
    Dim x_data_i As String
    Dim x_data_f As String
    x_data_i = "date(" & Format(msk_data_i, "yyyy,mm,dd") & ")"
    x_data_f = "date(" & Format(msk_data_f, "yyyy,mm,dd") & ")"
    If optOrdem(0) Then
        cr_relato.SortFields(0) = "+{Cheque_Folha.Data}"
        cr_relato.SortFields(1) = "+{Cheque_Folha.Numero}"
    Else
        cr_relato.SortFields(0) = "+{Cheque_Folha.Numero}"
        cr_relato.SortFields(1) = "+{Cheque_Folha.Data}"
    End If
    cr_relato.Formulas(0) = "f_empresa = BeforeReadingRecords;""" & g_nome_empresa & """"
    cr_relato.Formulas(1) = "f_data_hoje =  BeforeReadingRecords;""" & msk_data & """"
    cr_relato.Formulas(2) = "f_data_inicial =  BeforeReadingRecords;""" & msk_data_i & """"
    cr_relato.Formulas(3) = "f_data_final =  BeforeReadingRecords;""" & msk_data_f & """"
    If optSituacao(0) Then
        cr_relato.Formulas(4) = "f_situacao =  BeforeReadingRecords;""" & "Cheques em Aberto" & """"
    ElseIf optSituacao(1) Then
        cr_relato.Formulas(4) = "f_situacao =  BeforeReadingRecords;""" & "Cheques Baixados" & """"
    ElseIf optSituacao(2) Then
        cr_relato.Formulas(4) = "f_situacao =  BeforeReadingRecords;""" & "Cheques Cancelados" & """"
    ElseIf optSituacao(3) Then
        cr_relato.Formulas(4) = "f_situacao =  BeforeReadingRecords;""" & "Todos os Cheques" & """"
    End If
    cr_relato.Formulas(5) = "f_cheque_posse =  BeforeReadingRecords;""" & Mid(cboChequePosse.Text, 3, Len(cboChequePosse.Text) - 2) & """"
    l_sql = "{Cheque_Folha.Empresa} = " & g_empresa
    l_sql = l_sql & " And {Cheque_Folha.Data} >= " & x_data_i
    l_sql = l_sql & " And {Cheque_Folha.Data} <= " & x_data_f
    If optSituacao(0) Then
        l_sql = l_sql & " And {Cheque_Folha.Situacao} = " & Chr(34) & "A" & Chr(34)
    ElseIf optSituacao(1) Then
        l_sql = l_sql & " And {Cheque_Folha.Situacao} = " & Chr(34) & "B" & Chr(34)
    ElseIf optSituacao(2) Then
        l_sql = l_sql & " And {Cheque_Folha.Situacao} = " & Chr(34) & "C" & Chr(34)
    End If
    If cboChequePosse.ListIndex > 0 Then
        l_sql = l_sql & " And {Cheque_Folha.Cheque em Posse} = " & cboChequePosse.ListIndex
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
        optSituacao(3).SetFocus
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
    optSituacao(3).SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        optSituacao(3).SetFocus
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
        MsgBox "Data final deve ser maior que a data inicial.", 64, "Atenção!"
        msk_data_f.SetFocus
    ElseIf Not optSituacao(0) And Not optSituacao(1) And Not optSituacao(2) And Not optSituacao(3) Then
        MsgBox "Escolha uma situação.", 64, "Atenção!"
        optSituacao(0).SetFocus
    ElseIf Not optOrdem(0) And Not optOrdem(1) Then
        MsgBox "Escolha uma ordem.", 64, "Atenção!"
        optOrdem(0).SetFocus
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
    Set tbl_cheque = bd_sgp.OpenTable("Cheque_Folha")
    PreencheCboChequePosse
    
    optSituacao(0) = True
    optOrdem(0) = True
    cboChequePosse.ListIndex = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        optSituacao(3).SetFocus
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
Private Sub optOrdem_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub optSituacao_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
