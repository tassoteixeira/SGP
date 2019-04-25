VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form emissao_suspencao 
   Caption         =   "Emissão de Suspenção"
   ClientHeight    =   5595
   ClientLeft      =   4140
   ClientTop       =   2070
   ClientWidth     =   7530
   Icon            =   "lst_suspencao.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5595
   ScaleWidth      =   7530
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5400
      Picture         =   "lst_suspencao.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4620
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3360
      Picture         =   "lst_suspencao.frx":199C
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprime Suspenção de Funcionário."
      Top             =   4620
      Width           =   795
   End
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1320
      Picture         =   "lst_suspencao.frx":2FA6
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Visualiza Suspenção de Funcionário."
      Top             =   4620
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7275
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   1320
         Picture         =   "lst_suspencao.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txt_testemunha 
         Height          =   315
         Index           =   1
         Left            =   180
         MaxLength       =   40
         TabIndex        =   15
         Top             =   3960
         Width           =   6915
      End
      Begin VB.TextBox txt_testemunha 
         Height          =   315
         Index           =   0
         Left            =   180
         MaxLength       =   40
         TabIndex        =   14
         Top             =   3600
         Width           =   6915
      End
      Begin VB.TextBox txt_motivo 
         Height          =   315
         Index           =   3
         Left            =   180
         TabIndex        =   12
         Top             =   2880
         Width           =   6915
      End
      Begin VB.TextBox txt_motivo 
         Height          =   315
         Index           =   2
         Left            =   180
         TabIndex        =   11
         Top             =   2520
         Width           =   6915
      End
      Begin VB.TextBox txt_motivo 
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   2160
         Width           =   6915
      End
      Begin VB.TextBox txt_motivo 
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   1800
         Width           =   6915
      End
      Begin VB.TextBox txt_dias_suspencao 
         Height          =   285
         Left            =   180
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1140
         Width           =   435
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc adodcFuncionario 
         Height          =   330
         Left            =   3540
         Top             =   540
         Visible         =   0   'False
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adodcFuncionario"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo dtcboFuncionario 
         Bindings        =   "lst_suspencao.frx":599A
         Height          =   315
         Left            =   2100
         TabIndex        =   5
         Top             =   480
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboFuncionario"
      End
      Begin VB.Label Label1 
         Caption         =   "&Testemunhas"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   13
         Top             =   3360
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "&Motivos da Suspenção"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Dias de S&uspenção"
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "&Funcionário"
         Height          =   255
         Index           =   9
         Left            =   2100
         TabIndex        =   4
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "&Data de Emissão"
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin Crystal.CrystalReport cr_relato 
      Left            =   120
      Top             =   4620
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\VB5\Sgp\Data\lst_suspencao.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_suspencao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_sql As String
Dim l_telefone As String
Dim lCodigo As Integer
Dim lData As Date

Private Empresa As New cEmpresa
Private Funcionario As New cFuncionario
Private MovAdvertenciaSuspencao As New cMovAdvertenciaSuspencao
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Empresa = Nothing
    Set Funcionario = Nothing
    Set MovAdvertenciaSuspencao = Nothing
End Sub
Private Sub Relatorio()
    Dim x_data As String
    Dim x_endereco As String
    
    If bdAccess Then
        cr_relato.ReportFileName = "\VB5\Sgp\Data\lst_suspencao.rpt"
    ElseIf bdSqlServer Then
        cr_relato.ReportFileName = "\VB5\Sgp\Data\SQL_lst_suspencao.rpt"
    End If
    x_data = g_cidade_empresa & ", "
    x_data = x_data & Format(Day(msk_data.Text), "00")
    x_data = x_data & " de " & Format((msk_data.Text), "mmmm")
    x_data = x_data & " de " & Year(msk_data.Text) & "."
    x_endereco = Trim(Empresa.Endereco)
    x_endereco = x_endereco & " - " & Trim(Empresa.Bairro)
    x_endereco = x_endereco & " - " & Trim(Empresa.Cidade)
    x_endereco = x_endereco & " - " & Trim(Empresa.Estado) & "."
    cr_relato.Formulas(0) = "f_empresa = BeforeReadingRecords;""" & g_nome_empresa & """"
    cr_relato.Formulas(1) = "f_cgc = BeforeReadingRecords;""" & Empresa.CGC & """"
    cr_relato.Formulas(2) = "f_inscricao = BeforeReadingRecords;""" & Empresa.InscricaoEstadual & """"
    cr_relato.Formulas(3) = "f_endereco = BeforeReadingRecords;""" & x_endereco & """"
    cr_relato.Formulas(4) = "f_dias =  BeforeReadingRecords;""" & Val(txt_dias_suspencao.Text) & """"
    cr_relato.Formulas(5) = "f_motivo_1 =  BeforeReadingRecords;""" & txt_motivo(0).Text & """"
    cr_relato.Formulas(6) = "f_motivo_2 =  BeforeReadingRecords;""" & txt_motivo(1).Text & """"
    cr_relato.Formulas(7) = "f_motivo_3 =  BeforeReadingRecords;""" & txt_motivo(2).Text & """"
    cr_relato.Formulas(8) = "f_motivo_4 =  BeforeReadingRecords;""" & txt_motivo(3).Text & """"
    cr_relato.Formulas(9) = "f_data =  BeforeReadingRecords;""" & x_data & """"
    cr_relato.Formulas(10) = "f_testemunha_1 =  BeforeReadingRecords;""" & txt_testemunha(0).Text & """"
    cr_relato.Formulas(11) = "f_testemunha_2 =  BeforeReadingRecords;""" & txt_testemunha(1).Text & """"
    l_sql = "{Funcionario.Empresa} = " & g_empresa
    l_sql = l_sql & " and {Funcionario.Codigo} = " & Val(dtcboFuncionario.BoundText)
    cr_relato.SelectionFormula = l_sql
    If bdSqlServer Then
        cr_relato.Connect = "DSN=sgp_data;UID=sa;PWD=" & gSenhaBD
        cr_relato.Password = gSenhaBD
    End If
    cr_relato.Action = 1
    If Not MovAdvertenciaSuspencao.LocalizarCodigo(g_empresa, CDate(msk_data.Text), Val(dtcboFuncionario.BoundText)) Then
        AtualTabeAdvertenciaSuspencao
        If MovAdvertenciaSuspencao.Incluir Then
            lCodigo = MovAdvertenciaSuspencao.CodigoFuncionario
            lData = MovAdvertenciaSuspencao.Data
        Else
            MsgBox "Erro ao incluir Advertencia/Suspenção", vbOKOnly + vbCritical, "Erro de Integridade!"
        End If
    Else
        AtualTabeAdvertenciaSuspencao
        If MovAdvertenciaSuspencao.Alterar(g_empresa, lData, lCodigo) Then
            lCodigo = MovAdvertenciaSuspencao.CodigoFuncionario
            lData = MovAdvertenciaSuspencao.Data
        Else
            MsgBox "Erro ao alterar Advertencia/Suspenção", vbOKOnly + vbCritical, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub AtualTabeAdvertenciaSuspencao()
    MovAdvertenciaSuspencao.Empresa = g_empresa
    MovAdvertenciaSuspencao.Data = msk_data.Text
    MovAdvertenciaSuspencao.CodigoFuncionario = Val(dtcboFuncionario.BoundText)
    MovAdvertenciaSuspencao.AdvertenciaouSuspencao = "S"
    MovAdvertenciaSuspencao.dia = Val(txt_dias_suspencao.Text)
    MovAdvertenciaSuspencao.Motivo1 = txt_motivo(0).Text
    MovAdvertenciaSuspencao.Motivo2 = txt_motivo(1).Text
    MovAdvertenciaSuspencao.Motivo3 = txt_motivo(2).Text
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data.Text = RetiraGString(1)
    dtcboFuncionario.SetFocus
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
    ElseIf Val(dtcboFuncionario.BoundText) = 0 Then
        MsgBox "Selecione um funcionario.", vbInformation, "Atenção!"
        dtcboFuncionario.SetFocus
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
        txt_dias_suspencao.SetFocus
    End If
End Sub
Private Sub dbcbo_funcionario_LostFocus()
    tbl_funcionario.Seek "=", g_empresa, dtcboFuncionario.BoundText
    If Not tbl_funcionario.NoMatch Then
        l_telefone = tbl_funcionario!Telefone
    End If
End Sub
Private Sub dtcboFuncionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_motivo(0).SetFocus
    End If
End Sub
Private Sub dtcboFuncionario_LostFocus()
    If dtcboFuncionario.BoundText <> "" Then
        If Funcionario.LocalizarCodigo(g_empresa, Val(dtcboFuncionario.BoundText)) Then
        Else
            MsgBox "Funcionário não cadastrado!", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    Set adodcFuncionario.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Funcionario WHERE Empresa = " & g_empresa & " And Situacao = " & preparaTexto("A") & " ORDER BY Nome")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        dtcboFuncionario.SetFocus
    End If
    If Not Empresa.LocalizarCodigo(g_empresa) Then
        MsgBox "Empresa não cadastrada!", vbInformation, "Atenção"
        Finaliza
        Unload Me
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
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboFuncionario.SetFocus
    End If
End Sub
Private Sub txt_dias_suspencao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_motivo(0).SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_motivo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index < 3 Then
            txt_motivo(Index + 1).SetFocus
        Else
            txt_testemunha(0).SetFocus
        End If
    End If
End Sub
Private Sub txt_testemunha_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index < 1 Then
            txt_testemunha(Index + 1).SetFocus
        Else
            cmd_visualizar.SetFocus
        End If
    End If
End Sub
