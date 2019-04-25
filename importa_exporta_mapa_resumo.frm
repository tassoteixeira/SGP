VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form importa_exporta_mapa_resumo 
   Caption         =   "Importa / Exporta dados do Mapa Resumo"
   ClientHeight    =   2310
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   6495
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "importa_exporta_mapa_resumo.frx":0000
   ScaleHeight     =   2310
   ScaleWidth      =   6495
   Begin VB.Frame frmDados 
      Height          =   1275
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.OptionButton optExporta 
         Caption         =   "Exporta Dados"
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   870
         Value           =   -1  'True
         Width           =   2595
      End
      Begin VB.OptionButton optImporta 
         Caption         =   "Importa Dados"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   870
         Width           =   2595
      End
      Begin MSMask.MaskEdBox msk_data_inicial 
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   450
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data_final 
         Height          =   300
         Left            =   2760
         TabIndex        =   4
         Top             =   450
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "&Data inicial"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "D&ata final"
         Height          =   195
         Index           =   8
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   1680
      Picture         =   "importa_exporta_mapa_resumo.frx":0446
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Confirma o processamento."
      Top             =   1380
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4020
      Picture         =   "importa_exporta_mapa_resumo.frx":1720
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1380
      Width           =   795
   End
End
Attribute VB_Name = "importa_exporta_mapa_resumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lSQL As String
Private rsTabela As New adodb.Recordset
Private MovMapaResumo As New cMovimentoMapaResumo
Function CriaTabela(ByVal xStrCaminhoDB As String, ByVal xNomeTabela As String) As Boolean

'define as variáveis objeto a serem usadas
Dim catDB As ADOX.Catalog
Dim novaTabela As ADOX.Table
Dim i As Integer

On Error GoTo trata_erro
    
    CriaTabela = False
    Set catDB = New ADOX.Catalog
    
    'abre o objeto catalogo
    catDB.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & xStrCaminhoDB
    
    Set novaTabela = New ADOX.Table
    
    'cria um novo objeto table
    With novaTabela
        .name = xNomeTabela
        'cria campos e os anexa a coleção columns do novo objeto table
        With .Columns
            .Append "Empresa", adSmallInt
            .Append "Data", adDate
            .Append "Numero", adInteger
            .Append "ECF Numero", adSmallInt
            .Append "Contagem de Operacao Inicial", adInteger
            .Append "Contagem de Operacao Final", adInteger
            .Append "Totalizador Geral Final", adCurrency
            .Append "Totalizador Geral Inicial", adCurrency
            .Append "Cancelamento de Item", adCurrency
            .Append "Valor Contabil", adCurrency
            .Append "Isentas Nao Tributadas", adCurrency
            .Append "Substituicao Tributaria", adCurrency
            .Append "ICMS 17", adCurrency
            .Append "Contador de Reducoes Z", adInteger
            .Append "Observacao 1", adVarWChar, 50
            .Append "Observacao 2", adVarWChar, 50
        End With
    End With
    
    'cria a nova tabela incluindo o objeto table a coleção tables do banco de dados
    catDB.Tables.Append novaTabela
    
    'Define Observacao 1 e 2 como Allow Zero Lenght
    Set novaTabela = catDB.Tables(xNomeTabela)
    novaTabela.Columns("Observacao 1").Properties("Jet OLEDB:Allow Zero Length").Value = True
    novaTabela.Columns("Observacao 2").Properties("Jet OLEDB:Allow Zero Length").Value = True
    
    Set catDB = Nothing
    Set novaTabela = Nothing
    CriaTabela = True
    Exit Function

trata_erro:
    MsgBox Error
End Function
Private Sub Finaliza()
    Set MovMapaResumo = Nothing
End Sub
Private Sub Processamento()
    If optImporta.Value = True Then
        ProcessamentoImportaMapaResumo
    Else
        ProcessamentoExportaMapaResumo
    End If
End Sub
Private Sub ProcessamentoExportaMapaResumo()
    Dim cnnBDDisquete As New adodb.Connection
    Dim lRecordsAffected As Long
    Dim xFimLoop As Boolean
    Dim i As Long
    
    i = 0
    xFimLoop = False
    If CriaNovoMDB("A:\Mapa_Resumo.mdb", 4) Then
        If CriaTabela("A:\Mapa_Resumo.mdb", "Mapa_Resumo") Then
            If MovMapaResumo.LocalizarUltimo(g_empresa) Then
                cnnBDDisquete.Mode = adModeRead
                Set cnnBDDisquete = New adodb.Connection
                If bdAccess Then
                    cnnBDDisquete.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=A:\Mapa_Resumo.Mdb"
                    cnnBDDisquete.Open '"Driver={Microsoft Access Driver (*.mdb)};Dbq=A:\Mapa_Resumo.Mdb;Uid=Admin;Pwd=;"
                End If
                If Not MovMapaResumo.LocalizarData(g_empresa, CDate(msk_data_inicial.Text)) Then
                    Call MovMapaResumo.LocalizarPrimeiro
                End If
                Do Until xFimLoop = True
                    If MovMapaResumo.Data >= CDate(msk_data_inicial.Text) And MovMapaResumo.Data <= CDate(msk_data_final.Text) Then
                        gSQL = "INSERT INTO Mapa_Resumo ( Empresa, Data, Numero, [ECF Numero], [Contagem de Operacao Inicial], "
                        gSQL = gSQL & "[Contagem de Operacao Final], [Totalizador Geral Final], [Totalizador Geral Inicial], [Cancelamento de Item], [Valor Contabil], "
                        gSQL = gSQL & "[Isentas Nao Tributadas], [Substituicao Tributaria], [ICMS 17], [Contador de Reducoes Z], [Observacao 1], "
                        gSQL = gSQL & "[Observacao 2] ) VALUES ( "
                        Call sqlNumero(1, MovMapaResumo.Empresa, ", ")
                        Call sqlData(1, MovMapaResumo.Data, ", ")
                        Call sqlNumero(1, MovMapaResumo.numero, ", ")
                        Call sqlNumero(1, MovMapaResumo.ECFNumero, ", ")
                        Call sqlNumero(1, MovMapaResumo.ContagemOperacaoInicial, ", ")
                        Call sqlNumero(1, MovMapaResumo.ContagemOperacaoFinal, ", ")
                        Call sqlValor(1, MovMapaResumo.TotalizadorGeralFinal, ", ")
                        Call sqlValor(1, MovMapaResumo.TotalizadorGeralInicial, ", ")
                        Call sqlValor(1, MovMapaResumo.CancelamentoItem, ", ")
                        Call sqlValor(1, MovMapaResumo.ValorContabil, ", ")
                        Call sqlValor(1, MovMapaResumo.Isentas, ", ")
                        Call sqlValor(1, MovMapaResumo.SubstituicaoTributaria, ", ")
                        Call sqlValor(1, MovMapaResumo.ICMS17, ", ")
                        Call sqlNumero(1, MovMapaResumo.ContadorReducoesZ, ", ")
                        Call sqlTexto(1, MovMapaResumo.Observacao1, ", ")
                        Call sqlTexto(1, MovMapaResumo.Observacao2, " )")
                        cnnBDDisquete.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
                        If Not lRecordsAffected > 0 Then
                            MsgBox "Registro não foi gravado!", vbInformation, "Duplicidade de Registro!"
                        Else
                            i = i + lRecordsAffected
                        End If
                    End If
                    If Not MovMapaResumo.LocalizarProximo Then
                        Exit Do
                    End If
                Loop
                cnnBDDisquete.Close
                Set cnnBDDisquete = Nothing
            End If
        End If
    End If
    If i > 0 Then
        MsgBox "Foram exportados para o disquete " & i & " registros!", vbInformation, "Processamento Concluído!"
    End If
End Sub
Private Sub ProcessamentoImportaMapaResumo()
    Dim cnnBDDisquete As New adodb.Connection
    Dim i As Long
    
    i = 0
    cnnBDDisquete.Mode = adModeRead
    Set cnnBDDisquete = New adodb.Connection
    If bdAccess Then
        cnnBDDisquete.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=A:\Mapa_Resumo.Mdb"
        cnnBDDisquete.Open '"Driver={Microsoft Access Driver (*.mdb)};Dbq=A:\Mapa_Resumo.Mdb;Uid=Admin;Pwd=;"
    End If

    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT * "
    lSQL = lSQL & "  FROM Mapa_Resumo"
    lSQL = lSQL & " WHERE Data >= #" & Format(msk_data_inicial.Text, "mm/dd/yyyy") & "#"
    lSQL = lSQL & "   AND Data <= #" & Format(msk_data_final.Text, "mm/dd/yyyy") & "#"
    
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    rsTabela.CursorLocation = adUseClient
    rsTabela.Open lSQL, cnnBDDisquete.ConnectionString, adOpenForwardOnly, adLockReadOnly
    
    'Verifica Registros
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            MovMapaResumo.Empresa = rsTabela("Empresa").Value
            MovMapaResumo.Data = rsTabela("Data").Value
            MovMapaResumo.numero = rsTabela("Numero").Value
            MovMapaResumo.ECFNumero = rsTabela("ECF Numero").Value
            MovMapaResumo.ContagemOperacaoInicial = rsTabela("Contagem de Operacao Inicial").Value
            MovMapaResumo.ContagemOperacaoFinal = rsTabela("Contagem de Operacao Final").Value
            MovMapaResumo.TotalizadorGeralFinal = rsTabela("Totalizador Geral Final").Value
            MovMapaResumo.TotalizadorGeralInicial = rsTabela("Totalizador Geral Inicial").Value
            MovMapaResumo.CancelamentoItem = rsTabela("Cancelamento de Item").Value
            MovMapaResumo.ValorContabil = rsTabela("Valor Contabil").Value
            MovMapaResumo.Isentas = rsTabela("Isentas").Value
            MovMapaResumo.SubstituicaoTributaria = rsTabela("Substituicao Tributaria").Value
            MovMapaResumo.ICMS17 = rsTabela("ICMS 17").Value
            MovMapaResumo.ContadorReducoesZ = rsTabela("Contador de Reducoes Z").Value
            MovMapaResumo.Observacao1 = rsTabela("Observacao 1").Value
            MovMapaResumo.Observacao2 = rsTabela("Observacao 2").Value
            If Not MovMapaResumo.Incluir Then
                MsgBox "Registro não foi gravado!", vbInformation, "Duplicidade de Registro!"
            Else
                i = i + 1
            End If
            rsTabela.MoveNext
        Loop
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
    cnnBDDisquete.Close
    Set cnnBDDisquete = Nothing
    If i > 0 Then
        MsgBox "Foram importados do disquete " & i & " registros.", vbInformation, "Processamento Concluído!"
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        Processamento
        cmd_sair.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox "Operação não concluída!", vbInformation, "Erro no processamento!"
    Exit Sub
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data_inicial.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_inicial.SetFocus
    ElseIf Not IsDate(msk_data_final.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_final.SetFocus
    ElseIf Not IsDate(msk_data_final.Text) >= IsDate(msk_data_inicial.Text) Then
        MsgBox "A data final deve ser igual ou maior que " & msk_data_inicial & " .", vbInformation, "Atenção!"
        msk_data_final.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Screen.MousePointer = 1
End Sub
Private Sub Form_Load()
    CentraForm Me
    
    msk_data_inicial.Text = Format(g_data_def, "dd/mm/yyyy")
    msk_data_final.Text = Format(g_data_def, "dd/mm/yyyy")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_final_GotFocus()
    msk_data_final.SelStart = 0
    msk_data_final.SelLength = 5
End Sub
Private Sub msk_data_final_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub msk_data_inicial_GotFocus()
    msk_data_inicial.SelStart = 0
    msk_data_inicial.SelLength = 5
End Sub
Private Sub msk_data_inicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_final.SetFocus
    End If
End Sub
