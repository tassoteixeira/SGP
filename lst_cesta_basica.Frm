VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form emissao_cesta_basica 
   Caption         =   "Emissão de Recibo de Cesta Básica"
   ClientHeight    =   2715
   ClientLeft      =   75
   ClientTop       =   5040
   ClientWidth     =   6090
   Icon            =   "lst_cesta_basica.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2715
   ScaleWidth      =   6090
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   960
      Picture         =   "lst_cesta_basica.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Visualiza recibo de cesta básica."
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2640
      Picture         =   "lst_cesta_basica.frx":1A24
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprime recibo de cesta básica."
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4320
      Picture         =   "lst_cesta_basica.frx":302E
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1740
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5835
      Begin VB.TextBox txt_valor 
         Height          =   285
         Left            =   4560
         MaxLength       =   14
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   1320
         Picture         =   "lst_cesta_basica.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txt_funcionario 
         Height          =   300
         Left            =   180
         MaxLength       =   40
         TabIndex        =   8
         Top             =   840
         Width           =   5475
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
         Left            =   1680
         Top             =   1200
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
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
         Bindings        =   "lst_cesta_basica.frx":599A
         Height          =   315
         Left            =   180
         TabIndex        =   7
         Top             =   1140
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboFuncionario"
      End
      Begin VB.Label Label10 
         Caption         =   "V&alor"
         Height          =   285
         Index           =   16
         Left            =   4560
         TabIndex        =   4
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "&Funcionário"
         Height          =   255
         Index           =   9
         Left            =   180
         TabIndex        =   6
         Top             =   900
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
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\VB5\Sgp\Data\lst_cesta_basica.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_cesta_basica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_sql As String
Private Empresa As New cEmpresa
Private Funcionario As New cFuncionario
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Empresa = Nothing
    Set Funcionario = Nothing
End Sub
Private Sub PosicionaProximoRegistro()
'    dta_funcionario.Recordset.FindFirst "Codigo = " & Val(dtcboFuncionario.BoundText)
'    If Not dta_funcionario.Recordset.EOF Then
'        dta_funcionario.Recordset.MoveNext
'        If Not dta_funcionario.Recordset.EOF Then
'            dtcboFuncionario.BoundText = dta_funcionario.Recordset!Codigo
'        End If
'    End If
    dtcboFuncionario.SetFocus
End Sub
Private Sub Relatorio()
    Dim x_mes_ano As String
    Dim x_data As String
    Dim x_extenso As String
    
    If bdAccess Then
        cr_relato.ReportFileName = "\VB5\Sgp\Data\lst_cesta_basica.rpt"
    ElseIf bdSqlServer Then
        If UCase(g_nome_empresa) Like "*RATINHO*" Then
            cr_relato.ReportFileName = "\VB5\Sgp\Data\SQL_lst_cesta_basica_ratinho.rpt"
        Else
            cr_relato.ReportFileName = "\VB5\Sgp\Data\SQL_lst_cesta_basica.rpt"
        End If
    End If
    
    
    
    If Month(msk_data.Text) > 1 Then
        x_mes_ano = Format(CDate(10 & "/" & Month(msk_data.Text) - 1 & "/" & Year(msk_data.Text)), "mmmm") & " de " & Year(msk_data.Text) & "."
    Else
        x_mes_ano = Format(CDate(10 & "/" & 12 & "/" & Year(msk_data.Text)), "mmmm") & " de " & Year(msk_data.Text) - 1 & "."
    End If
    x_data = g_cidade_empresa & ", "
    x_data = x_data & Format(Day(msk_data.Text), "00")
    x_data = x_data & " de " & Format((msk_data.Text), "mmmm")
    x_data = x_data & " de " & Year(msk_data.Text) & "."
    cr_relato.Formulas(0) = "f_data = BeforeReadingRecords;""" & x_data & """"
    cr_relato.Formulas(1) = "f_funcionario = BeforeReadingRecords;""" & txt_funcionario & """"
    cr_relato.Formulas(2) = "f_mes_ano = BeforeReadingRecords;""" & x_mes_ano & """"
    If dtcboFuncionario.BoundText <> "" Then
        If Funcionario.LocalizarCodigo(g_empresa, Val(dtcboFuncionario.BoundText)) Then
            cr_relato.Formulas(3) = "f_cpf = BeforeReadingRecords;""" & Funcionario.CPF & """"
        Else
            cr_relato.Formulas(3) = "f_cpf = BeforeReadingRecords;""" & " " & """"
        End If
    Else
        cr_relato.Formulas(3) = "f_cpf = BeforeReadingRecords;""" & " " & """"
    End If
    cr_relato.Formulas(4) = "f_valor = BeforeReadingRecords;""" & txt_valor.Text & """"
    cr_relato.Formulas(5) = "f_valor_extenso = BeforeReadingRecords;""" & "(" & FazExtenso(fValidaValor(txt_valor.Text)) & ")" & """"
    If Empresa.LocalizarCodigo(g_empresa) Then
        cr_relato.Formulas(6) = "f_cgc = BeforeReadingRecords;" & preparaTexto(fMascaraCNPJ(Empresa.CGC))
        cr_relato.Formulas(7) = "f_nome_empresa = BeforeReadingRecords;" & preparaTexto(Empresa.Nome)
        cr_relato.Formulas(8) = "f_endereco_empresa = BeforeReadingRecords;" & preparaTexto(Empresa.Endereco)
        cr_relato.Formulas(9) = "f_bairro_empresa = BeforeReadingRecords;" & preparaTexto(Empresa.Bairro)
        cr_relato.Formulas(10) = "f_cidade_empresa = BeforeReadingRecords;" & preparaTexto(Empresa.Cidade)
        cr_relato.Formulas(11) = "f_uf_empresa = BeforeReadingRecords;" & preparaTexto(Empresa.Estado)
    End If
    l_sql = "{Empresas.Codigo} = " & g_empresa
'    l_sql = "{Funcionario.empresa} = " & g_empresa
'    l_sql = l_sql & " and {Funcionario.codigo} = " & Val(dtcboFuncionario.BoundText)
'    l_sql = ""
    PosicionaProximoRegistro
    cr_relato.SelectionFormula = l_sql
    If bdSqlServer Then
        cr_relato.Connect = "DSN=sgp_data;UID=sa;PWD=" & gSenhaBD
        cr_relato.Password = gSenhaBD
    End If
    cr_relato.Action = 1
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
            cmd_imprimir.SetFocus
        End If
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf txt_valor = "" Then
        MsgBox "Informe o valor.", vbInformation, "Atenção!"
        txt_valor.SetFocus
    ElseIf txt_funcionario = "" Then
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
            cmd_visualizar.SetFocus
        End If
    End If
End Sub
Private Sub dtcboFuncionario_GotFocus()
    txt_funcionario.Visible = False
End Sub
Private Sub dtcboFuncionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub dtcboFuncionario_LostFocus()
    txt_funcionario.Visible = True
    If dtcboFuncionario.BoundText <> "" Then
        txt_funcionario.Text = dtcboFuncionario.Text
        cmd_imprimir.SetFocus
    Else
        txt_funcionario.Text = ""
        txt_funcionario.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        Set adodcFuncionario.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Funcionario WHERE Empresa = " & g_empresa & " And Situacao = " & preparaTexto("A") & " ORDER BY Nome")
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        msk_data.Text = CDate("15/" & Month(msk_data.Text) & "/" & Year(msk_data.Text))
        txt_valor.Text = "60,26"
        txt_valor.SetFocus
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
        txt_valor.SetFocus
    End If
End Sub
Private Sub txt_funcionario_Click()
    dtcboFuncionario.SetFocus
End Sub
Private Sub txt_funcionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub txt_valor_GotFocus()
    txt_valor.SelStart = 0
    txt_valor.SelLength = Len(txt_valor.Text)
End Sub
Private Sub txt_valor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_funcionario.SetFocus
    End If
End Sub
Private Sub txt_valor_LostFocus()
    txt_valor.Text = Format(txt_valor.Text, "###,###,##0.00")
End Sub
