VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form emissao_nota_cliente 
   Caption         =   "Emissão das Notas p/ Cliente"
   ClientHeight    =   3825
   ClientLeft      =   2610
   ClientTop       =   2010
   ClientWidth     =   7950
   Icon            =   "lst_nota_cliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3825
   ScaleWidth      =   7950
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1440
      Picture         =   "lst_nota_cliente.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Visualiza relação de notas p/ cliente."
      Top             =   2880
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3540
      Picture         =   "lst_nota_cliente.frx":1A24
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Imprime relação de notas p/ cliente."
      Top             =   2880
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5640
      Picture         =   "lst_nota_cliente.frx":302E
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2880
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.CheckBox chkNaoMudarData 
         Caption         =   "Não &Mudar Data Automaticamente"
         Height          =   255
         Left            =   4380
         TabIndex        =   29
         Top             =   1680
         Width           =   3195
      End
      Begin VB.CheckBox chkNotaConferida 
         Caption         =   "&Notas Conferidas"
         Height          =   255
         Left            =   5280
         TabIndex        =   4
         Top             =   300
         Width           =   2235
      End
      Begin VB.CheckBox chkUnificaEmpresa 
         Height          =   255
         Left            =   2100
         TabIndex        =   25
         Top             =   2280
         Width           =   435
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   3240
         Picture         =   "lst_nota_cliente.frx":46C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   3240
         Picture         =   "lst_nota_cliente.frx":599A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   3240
         Picture         =   "lst_nota_cliente.frx":6C74
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1860
         Width           =   495
      End
      Begin VB.TextBox txt_vencimento 
         Height          =   285
         Index           =   8
         Left            =   4980
         MaxLength       =   2
         TabIndex        =   14
         Top             =   660
         Width           =   315
      End
      Begin VB.TextBox txt_vencimento 
         Height          =   285
         Index           =   7
         Left            =   4620
         MaxLength       =   2
         TabIndex        =   13
         Top             =   660
         Width           =   315
      End
      Begin VB.TextBox txt_vencimento 
         Height          =   285
         Index           =   6
         Left            =   4260
         MaxLength       =   2
         TabIndex        =   12
         Top             =   660
         Width           =   315
      End
      Begin VB.TextBox txt_vencimento 
         Height          =   285
         Index           =   5
         Left            =   3900
         MaxLength       =   2
         TabIndex        =   11
         Top             =   660
         Width           =   315
      End
      Begin VB.TextBox txt_vencimento 
         Height          =   285
         Index           =   4
         Left            =   3540
         MaxLength       =   2
         TabIndex        =   10
         Top             =   660
         Width           =   315
      End
      Begin VB.TextBox txt_vencimento 
         Height          =   285
         Index           =   3
         Left            =   3180
         MaxLength       =   2
         TabIndex        =   9
         Top             =   660
         Width           =   315
      End
      Begin VB.TextBox txt_vencimento 
         Height          =   285
         Index           =   2
         Left            =   2820
         MaxLength       =   2
         TabIndex        =   8
         Top             =   660
         Width           =   315
      End
      Begin VB.TextBox txt_vencimento 
         Height          =   285
         Index           =   1
         Left            =   2460
         MaxLength       =   2
         TabIndex        =   7
         Top             =   660
         Width           =   315
      End
      Begin VB.TextBox txt_vencimento 
         Height          =   285
         Index           =   0
         Left            =   2100
         MaxLength       =   2
         TabIndex        =   6
         Top             =   660
         Width           =   315
      End
      Begin VB.TextBox txt_cliente 
         Height          =   285
         Left            =   2100
         MaxLength       =   6
         TabIndex        =   16
         Top             =   1020
         Width           =   795
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   2100
         TabIndex        =   22
         Top             =   1860
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
         Left            =   2100
         TabIndex        =   19
         Top             =   1440
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
         Left            =   2100
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
      Begin MSAdodcLib.Adodc adodcCliente 
         Height          =   330
         Left            =   4380
         Top             =   1020
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
         Caption         =   "adodcCliente"
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
      Begin MSDataListLib.DataCombo dtcboCliente 
         Bindings        =   "lst_nota_cliente.frx":7F4E
         Height          =   315
         Left            =   2940
         TabIndex        =   17
         Top             =   1020
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Razao Social"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboCliente"
      End
      Begin VB.Label Label3 
         Caption         =   "&Unifica empresas"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "&Codigos dos vencimentos"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "C&liente"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   15
         Top             =   1020
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1500
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1935
      End
   End
   Begin Crystal.CrystalReport cr_relato 
      Left            =   180
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\VB5\Sgp\Data\lst_nota_cliente.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   660
      Top             =   3180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_nota_cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lSQL As String
Dim lMensagemCobranca As String
Dim lDataVencimento As Date

Private Cliente As New cCliente
Private Configuracao  As New cConfiguracao
Private Empresa As New cEmpresa
Private MovimentoNotaAbastecimento As New cMovimentoNotaAbastecimento

Dim rsCliente As adodb.Recordset

Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    FinalizaProcessoCaixa
    Set Cliente = Nothing
    Set Configuracao = Nothing
    Set Empresa = Nothing
    Set MovimentoNotaAbastecimento = Nothing
End Sub
Private Sub BuscaConfiguracao()
    lMensagemCobranca = ""
    If Configuracao.LocalizarCodigo(g_empresa) Then
        lMensagemCobranca = Configuracao.MensagemCobranca
    Else
        MsgBox "Não existe configuração para esta empresa.", vbInformation, "Ajustar Configuração do Sistema!"
    End If
End Sub
Private Sub Relatorio()
    Dim x_data_i As String
    Dim x_data_f As String
    Dim x_nome_empresa As String
    Dim xNotaConferida As Boolean
    
    If chkNotaConferida.Value = 0 Then
        xNotaConferida = False
    Else
        xNotaConferida = True
    End If
    If MovimentoNotaAbastecimento.TotalData(g_empresa, Cliente.Codigo, CDate(msk_data_i.Text), CDate(msk_data_f.Text), xNotaConferida) > 0 Then
        If Empresa.LocalizarCodigo(Cliente.Empresa) Then
            x_nome_empresa = Empresa.Nome
        Else
            x_nome_empresa = g_nome_empresa
        End If
        If bdAccess Then
            cr_relato.ReportFileName = "\VB5\Sgp\Data\lst_nota_cliente.rpt"
            x_data_i = "date(" & Format(msk_data_i.Text, "yyyy,mm,dd") & ")"
            x_data_f = "date(" & Format(msk_data_f.Text, "yyyy,mm,dd") & ")"
            x_nome_empresa = """" & x_nome_empresa & """"
        ElseIf bdSqlServer Then
            cr_relato.ReportFileName = "\VB5\Sgp\Data\SQL_lst_nota_cliente.rpt"
            x_data_i = "date(" & Format(msk_data_i.Text, "yyyy,MM,dd") & ")"
            x_data_f = "date(" & Format(msk_data_f.Text, "yyyy,MM,dd") & ")"
'            x_nome_empresa = preparaTexto(x_nome_empresa)
        End If
        
        lMensagemCobranca = "Vencimento em: " & lDataVencimento
        If chkNotaConferida.Value = 0 Then
            cr_relato.SortFields(0) = "+{Movimento_Nota_Abastecimento.Data do Abastecimento}"
        Else
            cr_relato.SortFields(0) = "+{Movimento_Nota_Abastecimento.Data da Conferencia}"
        End If
        cr_relato.SortFields(1) = "+{Movimento_Nota_Abastecimento.Numero da Nota}"
        cr_relato.Formulas(0) = "f_empresa = BeforeReadingRecords;""" & x_nome_empresa & """"
        cr_relato.Formulas(1) = "f_data_hoje =  BeforeReadingRecords;""" & msk_data & """"
        cr_relato.Formulas(2) = "f_data_inicial =  BeforeReadingRecords;""" & msk_data_i & """"
        cr_relato.Formulas(3) = "f_data_final =  BeforeReadingRecords;""" & msk_data_f & """"
        cr_relato.Formulas(4) = "f_cliente =  BeforeReadingRecords;""" & Cliente.RazaoSocial & """"
        cr_relato.Formulas(5) = "f_telefone =  BeforeReadingRecords;""" & fMascaraTelefone(Cliente.Telefone) & """"
        cr_relato.Formulas(6) = "f_mensagem_cobranca =  BeforeReadingRecords;""" & lMensagemCobranca & """"
        If chkNotaConferida.Value = 0 Then
            cr_relato.Formulas(7) = "f_conferencia =  BeforeReadingRecords;""" & "" & """"
        Else
            cr_relato.Formulas(7) = "f_conferencia =  BeforeReadingRecords;""" & "POR CONFERÊNCIA" & """"
        End If
        cr_relato.Formulas(8) = "f_nome_cidade =  BeforeReadingRecords;""" & g_cidade_empresa & "," & """"
        If chkUnificaEmpresa.Value = 0 Then
            lSQL = "{Movimento_Nota_Abastecimento.Empresa} = " & g_empresa & " AND "
        Else
            lSQL = ""
        End If
        lSQL = lSQL & "{Movimento_Nota_Abastecimento.Codigo do Cliente} = " & Cliente.Codigo
        If chkNotaConferida.Value = 0 Then
            lSQL = lSQL & " And {Movimento_Nota_Abastecimento.Data do Abastecimento} >= " & x_data_i
            lSQL = lSQL & " And {Movimento_Nota_Abastecimento.Data do Abastecimento} <= " & x_data_f
        Else
            lSQL = lSQL & " And {Movimento_Nota_Abastecimento.Data da Conferencia} >= " & x_data_i
            lSQL = lSQL & " And {Movimento_Nota_Abastecimento.Data da Conferencia} <= " & x_data_f
        End If
        
        If bdSqlServer Then
            cr_relato.Connect = "DSN=sgp_data;UID=sa;PWD=" & gSenhaBD
            cr_relato.Password = gSenhaBD
        End If
        cr_relato.SelectionFormula = lSQL
        If chkNaoMudarData.Value = 0 Then
            cr_relato.CopiesToPrinter = 2
        Else
            cr_relato.CopiesToPrinter = 1
        End If
        cr_relato.Action = 1
    Else
        MsgBox "O cliente " & Trim(Cliente.RazaoSocial) & "." & Chr(10) & "Não tem movimento no período " & msk_data_i & " à " & msk_data_f & ".", vbInformation, "Mensagem para o operador!"
    End If
End Sub
Private Sub chkUnificaEmpresa_KeyPress(KeyAscii As Integer)
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
        chkUnificaEmpresa.SetFocus
    Else
        msk_data = RetiraGString(1)
        txt_vencimento(0).SetFocus
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
    chkUnificaEmpresa.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        chkUnificaEmpresa.SetFocus
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
            LoopCliente
            dtcboCliente.SetFocus
        End If
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf Val(txt_vencimento(0)) > 0 Or Val(txt_vencimento(1)) > 0 Or Val(txt_vencimento(2)) > 0 Or Val(txt_vencimento(3)) > 0 Or Val(txt_vencimento(4)) > 0 Or Val(txt_vencimento(5)) > 0 Or Val(txt_vencimento(6)) > 0 Or Val(txt_vencimento(7)) > 0 Or Val(txt_vencimento(8)) > 0 Then
        ValidaCampos = True
        Exit Function
    ElseIf Not IsDate(msk_data_i) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f) < CDate(msk_data_i) Then
        MsgBox "Data final deve ser maior ou igual a " & CDate(msk_data_i) & ".", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf dtcboCliente.BoundText = "" Then
        MsgBox "Selecione um cliente.", vbInformation, "Atenção!"
        dtcboCliente.SetFocus
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
            LoopCliente
        End If
    End If
End Sub
Private Sub dtcboCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_i.SetFocus
    End If
End Sub
Private Sub dtcboCliente_LostFocus()
    If dtcboCliente.BoundText <> "" Then
        If Cliente.LocalizarCodigo(CLng(dtcboCliente.BoundText)) Then
            If Cliente.Codigo <> Val(txt_cliente.Text) Then
                txt_cliente.Text = Cliente.Codigo
            End If
            g_string = CalculaDataAbastecimentoVencimento(Cliente.CodigoVencimento, CDate(msk_data.Text))
            If g_string <> "" Then
                msk_data_i.Text = RetiraGString(1)
                msk_data_f.Text = RetiraGString(2)
                lDataVencimento = RetiraGString(3)
            End If
            g_string = ""
            cmd_imprimir.SetFocus
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        txt_vencimento(0).SetFocus
    End If
    Screen.MousePointer = 1
    
    If Len(g_string) > 0 Then
        If RetiraGString(1) = "ConfereNotaAbastecimento" Then
            txt_cliente.Text = RetiraGString(3)
            dtcboCliente.BoundText = CLng(RetiraGString(3))
            msk_data_i.Text = RetiraGString(4)
            msk_data_f.Text = RetiraGString(4)
            lDataVencimento = RetiraGString(5)
            chkNotaConferida.Value = 1
            If Cliente.LocalizarCodigo(CLng(txt_cliente.Text)) Then
            End If
            cmd_visualizar_Click
        End If
    End If
    
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
    Set adodcCliente.Recordset = Conectar.RsConexao("SELECT Codigo, [Razao Social] FROM Cliente WHERE Inativo = " & preparaBooleano(False) & " ORDER BY [Razao Social]")
    BuscaConfiguracao
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chkUnificaEmpresa.SetFocus
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
        txt_vencimento(0).SetFocus
    End If
End Sub
Private Sub txt_cliente_GotFocus()
    txt_cliente.SelStart = 0
    txt_cliente.SelLength = Len(txt_cliente.Text)
End Sub
Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboCliente.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_cliente_LostFocus()
    If Val(txt_cliente.Text) > 0 Then
        If Cliente.LocalizarCodigo(CLng(txt_cliente.Text)) Then
            If Cliente.Inativo = True Then
                MsgBox "O cliente " & Trim(Cliente.RazaoSocial) & " está inativo.", vbInformation, "Atenção!"
                dtcboCliente.BoundText = ""
                txt_cliente.SetFocus
                Exit Sub
            Else
                dtcboCliente.BoundText = CLng(txt_cliente.Text)
                msk_data_i.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Cliente não cadastrado.", vbInformation, "Atenção!"
            txt_cliente.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub LoopCliente()
    Dim i As Integer
    Dim xCodigoVencimento As String
    
    If Val(txt_vencimento(0).Text) = 0 And Val(txt_vencimento(1).Text) = 0 And Val(txt_vencimento(2).Text) = 0 And Val(txt_vencimento(3).Text) = 0 And Val(txt_vencimento(4).Text) = 0 And Val(txt_vencimento(5).Text) = 0 And Val(txt_vencimento(6).Text) = 0 And Val(txt_vencimento(7).Text) = 0 And Val(txt_vencimento(8).Text) = 0 Then
        Relatorio
    Else
        xCodigoVencimento = ""
        For i = 0 To 8
            If txt_vencimento(i).Text <> "" Then
                If Len(xCodigoVencimento) > 0 Then
                    xCodigoVencimento = xCodigoVencimento & ", "
                End If
                xCodigoVencimento = xCodigoVencimento & txt_vencimento(i)
            End If
        Next
        
        'Prepara SQL
        lSQL = ""
        lSQL = lSQL & "   SELECT Codigo, [Codigo do Vencimento]"
        lSQL = lSQL & "     FROM Cliente"
        lSQL = lSQL & "    WHERE Inativo = " & preparaBooleano(False)
        lSQL = lSQL & "      AND [Codigo do Vencimento] IN (" & xCodigoVencimento & ")"
        
        lSQL = lSQL & " ORDER BY [Razao Social]"
        'Abre RecordSet
        Set rsCliente = New adodb.Recordset
        Set rsCliente = Conectar.RsConexao(lSQL)
        
        If rsCliente.RecordCount > 0 Then
            rsCliente.MoveFirst
            MsgBox "Foi selecionado " & rsCliente.RecordCount & " cliente(s) para impressão." & vbCrLf & "Pode ser que nem todos tenha algo a ser impresso.", vbExclamation + vbOKOnly + vbInformation, "Mensagem ao Operador!"
            Do Until rsCliente.EOF
                If chkNaoMudarData.Value = 0 Then
                    g_string = CalculaDataAbastecimentoVencimento(rsCliente("Codigo do Vencimento").Value, CDate(msk_data.Text))
                    If g_string <> "" Then
                        msk_data_i.Text = RetiraGString(1)
                        msk_data_f.Text = RetiraGString(2)
                        lDataVencimento = RetiraGString(3)
                    End If
                Else
                    lDataVencimento = Format(Date, "dd/mm/yyyy")
                End If
                g_string = ""
                
                If Cliente.LocalizarCodigo(rsCliente("Codigo").Value) Then
                    Relatorio
                Else
                    MsgBox "Não foi possível localizar o cliente " & rsCliente("Codigo").Value, vbExclamation + vbOKOnly + vbCritical, "Erro de Integridade!"
                End If
                rsCliente.MoveNext
            Loop
        End If
        rsCliente.Close
        Set rsCliente = Nothing
    End If
End Sub
Private Sub txt_vencimento_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 0 And txt_vencimento(0) = "" Then
            txt_cliente.SetFocus
            Exit Sub
        End If
        If Index > 0 And txt_vencimento(Index) = "" Then
            cmd_imprimir.SetFocus
            Exit Sub
        End If
        SendKeys "{Tab}"
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_vencimento_LostFocus(Index As Integer)
    If Val(txt_vencimento(0)) = 0 And Val(txt_vencimento(1)) = 0 And Val(txt_vencimento(2)) = 0 And Val(txt_vencimento(3)) = 0 And Val(txt_vencimento(4)) = 0 And Val(txt_vencimento(5)) = 0 And Val(txt_vencimento(6)) = 0 And Val(txt_vencimento(7)) = 0 And Val(txt_vencimento(8)) = 0 Then
        cmd_visualizar.Enabled = True
    Else
        cmd_visualizar.Enabled = False
    End If
End Sub
