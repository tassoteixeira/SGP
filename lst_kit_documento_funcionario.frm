VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form emissao_kit_documento_funcionario 
   Caption         =   "Emissão do Kit de Documentos de Funcionário"
   ClientHeight    =   4455
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   7515
   Icon            =   "lst_kit_documento_funcionario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4455
   ScaleWidth      =   7515
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1980
      Picture         =   "lst_kit_documento_funcionario.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Visualiza o tipo de documento selecionado."
      Top             =   3480
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4740
      Picture         =   "lst_kit_documento_funcionario.frx":1A24
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3480
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7275
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_kit_documento_funcionario.frx":30B6
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_desconto 
         Height          =   555
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2160
         Width           =   5475
      End
      Begin VB.TextBox txt_servico 
         Height          =   555
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1500
         Width           =   5475
      End
      Begin VB.ComboBox cbo_tipo_documento 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   5475
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
      Begin MSMask.MaskEdBox msk_hora_f 
         Height          =   315
         Left            =   6480
         TabIndex        =   14
         Top             =   2820
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_hora_i 
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Top             =   2820
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc adodcFuncionario 
         Height          =   330
         Left            =   3180
         Top             =   720
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
         Bindings        =   "lst_kit_documento_funcionario.frx":4390
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Top             =   660
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
      Begin VB.Label Label2 
         Caption         =   "&Descontos"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Hora de Entrada"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   2820
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "Hora de Saida"
         Height          =   255
         Left            =   4920
         TabIndex        =   13
         Top             =   2820
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "&S&erviços"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "F&uncionário"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label7 
         Caption         =   "&Tipo do Documento"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_kit_documento_funcionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bd_sgp_tmp As Database

Private Empresa As New cEmpresa
Private Funcionario As New cFuncionario

Dim tbl_dados_funcionario As Table
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_dados_funcionario.Close
    
    Set Empresa = Nothing
    Set Funcionario = Nothing
End Sub
Private Sub RegravaRegistro()
    With tbl_dados_funcionario
        If Empresa.LocalizarCodigo(g_empresa) Then
            .Edit
            !Empresa = Empresa.Nome
            !Endereço1 = Empresa.Endereco
            !CGC = fMascaraCNPJ(Empresa.CGC)
            !Funcionario = Funcionario.Nome
            !CTPS = Funcionario.CarteiraTrabalho
            !Serie = Funcionario.SerieCarteiraTrabalho
            !Cargo = Funcionario.Cargo
            !Serviços = txt_servico.Text
            !Hora1 = msk_hora_i.Text
            !Hora2 = msk_hora_f.Text
            !Salario = Format(Funcionario.SalarioBase, "###,###,##0.00")
            !Extenso = " "
            If !Salario > 0 Then
                !Extenso = FazExtenso(fValidaValor2(!Salario))
            End If
            !Descontos = txt_desconto
            !Dias30 = CDate(msk_data) + 29
            !Dias90 = CDate(msk_data) + 89
            !dia = Format(msk_data, "dd")
            !mes = Format(msk_data, "mmmm")
            !ano = Format(msk_data, "yyyy")
            .Update
        End If
    End With
End Sub
Private Sub Relatorio()
    Dim x_documento As String
    Dim retval As Long
    x_documento = gDrive & "\VB5\SGP\DOC\"
    If cbo_tipo_documento = "Acordo de Compensação de Horas de Trabalho" Then
        x_documento = x_documento & "ACORDO_DE_COMPENSAÇÃO_DE_HORAS_DE_TRABALHO.doc"
    ElseIf cbo_tipo_documento = "Acordo de Prorrogação de Hora" Then
        x_documento = x_documento & "ACORDO_DE_PRORROGAÇAO_DE_HORA.doc"
    ElseIf cbo_tipo_documento = "Contrato de Trabalho de Experiência" Then
        x_documento = x_documento & "CONTRATO_DE_TRABALHO_A_TÍTULO_DE_EXPERIÊNCI1.doc"
    ElseIf cbo_tipo_documento = "Kit de Admissão de Funcionário" Then
        x_documento = x_documento & "KIT_DE_ADMISSÃO_DE_FUNCIONÁRIO.doc"
    ElseIf cbo_tipo_documento = "Opção de Desistência de Vale Transporte" Then
        x_documento = x_documento & "OPÇÃO_DE_DESISTÊNCIA_DE_VALE_TRANSPORTE.doc"
    ElseIf cbo_tipo_documento = "Recibo de Cheques Devolvidos" Then
        x_documento = x_documento & "recibo_de_cheques_devolvidos.doc"
    ElseIf cbo_tipo_documento = "Recibo de Entrega de E.P.I." Then
        x_documento = x_documento & "RECIBO_DE_ENTREGA_DE_EPI.doc"
    ElseIf cbo_tipo_documento = "Regulamento Interno do Posto" Then
        x_documento = x_documento & "REGULAMENTO_INTERNO_DO_POST1.doc"
    End If
    If gArqTxt.FileExists("C:\Arquivos de programas\Microsoft Office\Office12\WinWord.exe") Then
        retval = Shell("C:\Arquivos de programas\Microsoft Office\Office12\WinWord.exe " & x_documento)
    ElseIf gArqTxt.FileExists("C:\Arquivos de programas\Microsoft Office\Office11\WinWord.exe") Then
        retval = Shell("C:\Arquivos de programas\Microsoft Office\Office11\WinWord.exe " & x_documento)
    ElseIf gArqTxt.FileExists("C:\Arquivos de Programas\Microsoft Office\Office\Winword.exe") Then
        retval = Shell("C:\Arquivos de Programas\Microsoft Office\Office\Winword.exe " & x_documento)
    Else
        MsgBox "Word Não encontrado", vbCritical + vbOKOnly, "Erro de Localização!"
    End If
End Sub
Private Sub cbo_tipo_documento_GotFocus()
    SendMessageLong cbo_tipo_documento.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_tipo_documento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_servico.SetFocus
    End If
End Sub
Private Sub TabelaFuncionarioRefresh()
'    dta_funcionario.RecordSource = "Select * From Funcionario Where Empresa = " & g_empresa & " And Situacao = 'A' Order By Nome"
'    dta_funcionario.Refresh
    Set adodcFuncionario.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Funcionario WHERE Empresa = " & g_empresa & " And Situacao = " & preparaTexto("A") & " ORDER BY Nome")
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf dtcboFuncionario.BoundText = "" Then
        MsgBox "Selecione o funcionario.", vbInformation, "Atenção!"
        dtcboFuncionario.SetFocus
    ElseIf cbo_tipo_documento.ListIndex = -1 Then
        MsgBox "Selecione o tipo de documento.", vbInformation, "Atenção!"
        cbo_tipo_documento.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = RetiraGString(1)
    dtcboFuncionario.SetFocus
    g_string = " "
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    If ValidaCampos Then
        If SelecionaImpressoraHP(Me) Then
            Call GravaAuditoria(1, Me.name, 6, "")
            RegravaRegistro
            Relatorio
            dtcboFuncionario.SetFocus
        End If
    End If
End Sub
Private Sub dtcboFuncionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub dtcboFuncionario_LostFocus()
    If dtcboFuncionario.BoundText <> "" Then
        If Funcionario.LocalizarCodigo(g_empresa, Val(dtcboFuncionario.BoundText)) Then
            txt_servico.Text = "Lavador; Enxugador; Limpeza; Serviços Gerais; Troca de Óleo; Cobrança; Escritório; Noturno;"
            txt_desconto.Text = "Falta de Caixa; Cheques Devolvidos Fora das Normas; Vendas Por Cartão Devolvidas; Prejuizos Com Desatenção ou Devoluções de Produtos ou Serviços; Faltas Injustificadas;"
            msk_hora_i.Text = "08:00"
            msk_hora_f.Text = "18:00"
            If Funcionario.Periodo = 1 Then
                msk_hora_i.Text = "06:00"
                msk_hora_f.Text = "14:00"
            ElseIf Funcionario.Periodo = 2 Then
                msk_hora_i.Text = "14:00"
                msk_hora_f.Text = "22:00"
            ElseIf Funcionario.Periodo = 3 Then
                msk_hora_i.Text = "22:00"
                msk_hora_f.Text = "06:00"
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    TabelaFuncionarioRefresh
    PreencheCboTipoDocumento
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(Date, "dd/mm/yyyy")
        cbo_tipo_documento.ListIndex = 0
        dtcboFuncionario.BoundText = ""
        dtcboFuncionario.SetFocus
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_visualizar_Click
    ElseIf KeyCode = vbKeyF9 Then
        KeyCode = 0
        cmd_visualizar_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    
    Set bd_sgp_tmp = OpenDatabase("SGP_DATA_TMP.MDB")
    Set tbl_dados_funcionario = bd_sgp_tmp.OpenTable("Dados_Funcionario")
    tbl_dados_funcionario.Index = "id_codigo"
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
Private Sub PreencheCboTipoDocumento()
    cbo_tipo_documento.Clear
    cbo_tipo_documento.AddItem "Acordo de Compensação de Horas de Trabalho"
    cbo_tipo_documento.ItemData(cbo_tipo_documento.NewIndex) = 0
    cbo_tipo_documento.AddItem "Acordo de Prorrogação de Hora"
    cbo_tipo_documento.ItemData(cbo_tipo_documento.NewIndex) = 1
    cbo_tipo_documento.AddItem "Contrato de Trabalho de Experiência"
    cbo_tipo_documento.ItemData(cbo_tipo_documento.NewIndex) = 2
    cbo_tipo_documento.AddItem "Kit de Admissão de Funcionário"
    cbo_tipo_documento.ItemData(cbo_tipo_documento.NewIndex) = 3
    cbo_tipo_documento.AddItem "Opção de Desistência de Vale Transporte"
    cbo_tipo_documento.ItemData(cbo_tipo_documento.NewIndex) = 4
    cbo_tipo_documento.AddItem "Recibo de Cheques Devolvidos"
    cbo_tipo_documento.ItemData(cbo_tipo_documento.NewIndex) = 5
    cbo_tipo_documento.AddItem "Recibo de Entrega de E.P.I."
    cbo_tipo_documento.ItemData(cbo_tipo_documento.NewIndex) = 6
    cbo_tipo_documento.AddItem "Regulamento Interno do Posto"
    cbo_tipo_documento.ItemData(cbo_tipo_documento.NewIndex) = 7
End Sub
Private Sub msk_hora_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub msk_hora_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_hora_f.SetFocus
    End If
End Sub
Private Sub txt_desconto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_hora_i.SetFocus
    End If
End Sub
Private Sub txt_servico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_desconto.SetFocus
    End If
End Sub
