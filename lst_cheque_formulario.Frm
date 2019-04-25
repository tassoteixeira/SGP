VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form listagem_cheque_formulario 
   Caption         =   "Listagem dos Cheques (formulário)"
   ClientHeight    =   2715
   ClientLeft      =   3270
   ClientTop       =   4230
   ClientWidth     =   4335
   Icon            =   "lst_cheque_formulario.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_cheque_formulario.frx":030A
   ScaleHeight     =   2715
   ScaleWidth      =   4335
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   960
      Picture         =   "lst_cheque_formulario.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprime relação do cheque (formulário)."
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2580
      Picture         =   "lst_cheque_formulario.frx":195A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1740
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1515
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_cheque_formulario.frx":2FEC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_cheque_formulario.frx":42C6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2700
         Picture         =   "lst_cheque_formulario.frx":55A0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
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
      Begin VB.Label Label2 
         Caption         =   "Data Final"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1080
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
      Begin VB.Label Label5 
         Caption         =   "Data de Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
End
Attribute VB_Name = "listagem_cheque_formulario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim l_linha_tab As Integer
    Dim l_coluna_tab As Integer
    Dim l_linha As Integer
Private Sub ZeraVariaveis()
    l_coluna_tab = 0
    l_linha_tab = 0
    l_linha = 0
End Sub
Private Sub Finaliza()
    tbl_cheque_emitido.Close
End Sub
Private Sub ImprimeDados()
    Dim x_extenso As String
    Dim x_mes As String
    Dim x_data As Date
    Dim x2_data As String
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    Printer.FontName = "Arial"
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    Do Until tbl_cheque_emitido.EOF
        If tbl_cheque_emitido!Empresa <> g_empresa Then
            Exit Do
        End If
        If tbl_cheque_emitido!Data > msk_data_f Then
            Exit Do
        End If
        If l_linha = 2 Then
            ImprimeGrade l_linha
            Printer.NewPage
            l_linha = 0
        End If
        'teste
        Printer.PSet (0, 0)
        l_linha = l_linha + 1
        l_linha_tab = l_linha * 10 - 10 + 2 + l_linha_tab
        l_coluna_tab = 0
        Printer.CurrentX = l_coluna_tab + 7
        Printer.CurrentY = l_linha_tab + 2
        Printer.Print "Número do Cheque"
        Printer.CurrentX = l_coluna_tab + 11
        Printer.CurrentY = l_linha_tab + 2
        Printer.Print tbl_cheque_emitido!numero
        Printer.CurrentX = l_coluna_tab + 15
        Printer.CurrentY = l_linha_tab + 2
        Printer.Print "R$"
        Printer.CurrentX = l_coluna_tab + 15.5
        Printer.CurrentY = l_linha_tab + 2
        Printer.Print Format(tbl_cheque_emitido!valor, "###,##0.00")
        x_extenso = FazExtenso(tbl_cheque_emitido!valor)
        Printer.CurrentX = l_coluna_tab + 1
        Printer.CurrentY = l_linha_tab + 3
        Printer.Print x_extenso
        Printer.CurrentX = l_coluna_tab + 1
        Printer.CurrentY = l_linha_tab + 5
        Printer.Print tbl_cheque_emitido!Nome
        x_data = tbl_cheque_emitido!Data
        x2_data = tbl_cheque_emitido!Data
        FazExtensoMes x2_data, x_mes
        Printer.CurrentX = l_coluna_tab + 6
        Printer.CurrentY = l_linha_tab + 6
        Printer.Print "Goiânia,"
        Printer.CurrentX = l_coluna_tab + 7.5
        Printer.CurrentY = l_linha_tab + 6
        Printer.Print Day(tbl_cheque_emitido!Data)
        Printer.CurrentX = l_coluna_tab + 8.5
        Printer.CurrentY = l_linha_tab + 6
        Printer.Print "de"
        Printer.CurrentX = l_coluna_tab + 9.5
        Printer.CurrentY = l_linha_tab + 6
        Printer.Print x_mes
        Printer.CurrentX = l_coluna_tab + 12
        Printer.CurrentY = l_linha_tab + 6
        Printer.Print "de"
        Printer.CurrentX = l_coluna_tab + 13
        Printer.CurrentY = l_linha_tab + 6
        Printer.Print Year(tbl_cheque_emitido!Data)
        Printer.CurrentX = l_coluna_tab + 8
        Printer.CurrentY = l_linha_tab + 8
        Printer.Print g_nome_empresa
        tbl_cheque_emitido.MoveNext
    Loop
    ImprimeGrade l_linha
    Printer.EndDoc
End Sub
Private Sub ImprimeGrade(x_vezes As Integer)
    Dim i As Integer
    'Seleciona tamanho da fonte
    l_coluna_tab = 0
    l_linha_tab = 0
    Printer.FontSize = 14
    Printer.DrawWidth = 6
    For i = 1 To l_linha
        l_linha_tab = i * 10 - 10 + 2 + l_linha_tab
        Printer.CurrentX = l_coluna_tab + 3
        Printer.CurrentY = l_linha_tab + 1.1
        Printer.Print "CÓPIA DE CHEQUE EMITIDO (FORMULÁRIO)"
        Printer.Line (l_coluna_tab, l_linha_tab)-(l_coluna_tab + 20, l_linha_tab + 10), , B
    Next
    l_coluna_tab = 0
    l_linha_tab = 0
    Printer.FontSize = 10
    Printer.DrawWidth = 4
End Sub
Private Sub Relatorio()
    Dim i As Integer
    ZeraVariaveis
    tbl_cheque_emitido.Index = "id_data"
    tbl_cheque_emitido.Seek ">", g_empresa, msk_data_i, 0
    If Not tbl_cheque_emitido.NoMatch Then
        If tbl_cheque_emitido!Empresa = g_empresa Then
            ImprimeDados
        End If
    End If
    cmd_sair.SetFocus
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cmd_imprimir.SetFocus
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
    cmd_imprimir.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cmd_imprimir.SetFocus
    Else
        msk_data_i = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    If ValidaCampos Then
        If SelecionaImpressoraHP(Me) Then
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
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    If Not IsDate(msk_data) Then
        msk_data = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i.SetFocus
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_imprimir_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    Set tbl_cheque_emitido = bd_sgp.OpenTable("Cheque_Emitidos")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
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
