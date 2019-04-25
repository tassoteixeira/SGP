VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form envia_email 
   Caption         =   "Envia Email"
   ClientHeight    =   6690
   ClientLeft      =   75
   ClientTop       =   4020
   ClientWidth     =   6990
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "envia_email.frx":0000
   ScaleHeight     =   6690
   ScaleWidth      =   6990
   Begin VB.CommandButton cmd_le_xml 
      Caption         =   "Lê XML"
      Height          =   255
      Left            =   3900
      TabIndex        =   13
      Top             =   6420
      Width           =   1995
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   1980
      Top             =   5940
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   240
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4800
      Picture         =   "envia_email.frx":0446
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5520
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   5235
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6735
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   1680
         TabIndex        =   12
         Top             =   2640
         Width           =   4935
      End
      Begin VB.TextBox txt_destinatario 
         Height          =   285
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   7
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox txt_assunto 
         Height          =   285
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   6
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox txt_anexo 
         Height          =   285
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   5
         Top             =   2220
         Width           =   4935
      End
      Begin VB.TextBox txt_mensagem 
         Height          =   1155
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   960
         Width           =   4935
      End
      Begin VB.Label Label4 
         Caption         =   "Destinatário"
         Height          =   255
         Left            =   60
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Assunto"
         Height          =   255
         Left            =   60
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Anexo(s)"
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   2220
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Mensagem"
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Enviar"
      Height          =   855
      Left            =   2820
      Picture         =   "envia_email.frx":1720
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancela o registro atual."
      Top             =   5520
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Conectar"
      Height          =   855
      Left            =   1020
      Picture         =   "envia_email.frx":29FA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Confirma o registro atual."
      Top             =   5520
      Width           =   795
   End
End
Attribute VB_Name = "envia_email"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lSQl As String
Dim lArqTxt As New FileSystemObject
Dim rst As New adodb.Recordset
Dim lTabela As String

Private Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Long
Private Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Private Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2
Private Sub AtualizaRecordset(xMaxRecords As Integer)
    On Error GoTo FileError
    rst.CursorLocation = adUseClient
    rst.MaxRecords = xMaxRecords
    rst.Open lSQl, cnnSGP, adOpenForwardOnly, adLockReadOnly
    Exit Sub
FileError:
    rst.Close
    rst.CursorLocation = adUseClient
    rst.MaxRecords = xMaxRecords
    rst.Open lSQl, cnnSGP, adOpenForwardOnly, adLockReadOnly
    Exit Sub
End Sub
Function BuscaDados() As Boolean
    BuscaDados = False
    lSQl = PreparaSQL("", "ORDER BY Codigo DESC")
    Call AtualizaRecordset(1)
    If Not rst.EOF Then
        'AtualTela
        BuscaDados = True
    End If
    rst.Close
End Function
Function BuscaRegistro(x_codigo As Integer) As Boolean
    BuscaRegistro = False
    On Error GoTo FileError
'    lSQL = PreparaSQL(" AND Codigo = " & lCodigo, "")
    Call AtualizaRecordset(1)
    If Not rst.EOF Then
        BuscaRegistro = True
'        AtualTela
    End If
    rst.Close
    Exit Function
FileError:
    rst.Close
    Exit Function
End Function
Private Sub Finaliza()
    'tbl_bomba.Close
    'tbl_combustivel.Close
End Sub
Private Sub cmd_cancelar_Click()
    Dim i As Long
    
    
    'Cria Contas_Pagar.XML
    If lArqTxt.FileExists(gDrive & gDiretorioData & lTabela & ".XML") Then
        lArqTxt.DeleteFile (gDrive & gDiretorioData & lTabela & ".XML")
    End If
    lSQl = "Select * From " & lTabela
    rst.Open lSQl, cnnSGP, adOpenForwardOnly, adLockReadOnly
    rst.save gDrive & gDiretorioData & lTabela & ".XML", adPersistXML
    rst.Close
    
    
    
    
    
    MAPISession1.UserName = "tasso_cerrado@uol.com.br"
    MAPISession1.Password = "lara28"
    
    'Inicia Recebimento de Email
    MAPISession1.SignOn
    MAPIMessages1.FetchUnreadOnly = True
    MAPIMessages1.SessionID = MAPISession1.SessionID
    MAPIMessages1.Fetch
    MsgBox " Você tem " & MAPIMessages1.MsgCount & " Mensagens..."
    MAPISession1.SignOff
    'Finaliza Recebimento de Email




    MAPISession1.DownLoadMail = True
    MAPISession1.SignOn
    
    MAPIMessages1.SessionID = MAPISession1.SessionID
    MAPIMessages1.Fetch
    List1.Clear
    If MAPIMessages1.MsgCount > 0 Then
        For i = 0 To MAPIMessages1.MsgCount - 1
            MAPIMessages1.MsgIndex = i
            List1.AddItem "De: " & MAPIMessages1.MsgOrigDisplayName & " Assunto: " & MAPIMessages1.MsgSubject
        Next
    End If
    MAPISession1.SignOff
    
    
    
    
    
    
    
    
    
    
      MAPISession1.SignOn
    
      MAPIMessages1.SessionID = MAPISession1.SessionID
    
      MAPIMessages1.Compose
      MAPIMessages1.RecipAddress = txt_destinatario.Text
      MAPIMessages1.MsgSubject = txt_assunto.Text
      MAPIMessages1.MsgNoteText = txt_mensagem.Text
    
      'anexa no final da mensagem
      MAPIMessages1.AttachmentPosition = Len(MAPIMessages1.MsgNoteText)
      'define o tipo de dados do anexo
      MAPIMessages1.AttachmentType = mapData
      'da um nome ao anexo
      MAPIMessages1.AttachmentName = "Anexos"
      'define o caminho e nome do arquivo a anexar
      MAPIMessages1.AttachmentPathName = txt_anexo.Text
    
      'envia o arquivo
      MAPIMessages1.send False
    
      MAPISession1.SignOff
End Sub
Private Sub cmd_le_xml_Click()
    'Lê ContasPagar.XML
    Dim i As Integer
    Dim catSchema As ADOX.Catalog
    Dim tblSchema As ADOX.Table
    Dim colSchema As ADOX.Column
    
    Set catSchema = New ADOX.Catalog
    Set catSchema.ActiveConnection = cnnSGP
    
    For Each tblSchema In catSchema.Tables
        If tblSchema.Type = "TABLE" Then
            If UCase(tblSchema.name) = UCase(lTabela) Then
                lSQl = "INSERT INTO " & tblSchema.name & " ( "
                i = 0
                For Each colSchema In tblSchema.Columns
                    i = i + 1
                    If i > 1 Then
                        lSQl = lSQl & ", [" & colSchema.name & "]"
                    Else
                        lSQl = lSQl & "[" & colSchema.name & "]"
                    End If
                Next
                lSQl = lSQl & " ) VALUES ( "
                Exit For
            End If
        End If
    Next
    If (MsgBox("Deseja limpar a tabela: " & lTabela & "?", vbDefaultButton2 + vbQuestion + vbYesNo, "Limpa dados")) = 6 Then
        cnnSGP.Execute "DELETE * FROM " & lTabela
    End If
    rst.Open gDrive & gDiretorioData & lTabela & ".XML", "Provider=MSPersist"
    rst.MoveFirst
    Do Until rst.EOF
        gSQL = lSQl
        For Each colSchema In tblSchema.Columns
            If colSchema.Type = adSmallInt Or colSchema.Type = adInteger Then
                Call sqlNumero(1, rst.Fields(colSchema.name), ", ")
            ElseIf colSchema.Type = adVarWChar Or colSchema.Type = adChar Or colSchema.Type = adVarChar Then
                Call sqlTexto(1, rst.Fields(colSchema.name), ", ")
            ElseIf colSchema.Type = adDBTimeStamp Then
                Call sqlData(1, rst.Fields(colSchema.name), ", ")
            ElseIf colSchema.Type = adNumeric Then
                Call sqlValor(1, rst.Fields(colSchema.name), ", ")
            Else
                MsgBox "Tipo de campo desconhecido"
                MsgBox colSchema.name
                MsgBox colSchema.Type
            End If
        Next
        i = Len(gSQL)
        Mid(gSQL, i - 1, 2) = " )"
        'grava registro
        cnnSGP.Execute gSQL
        rst.MoveNext
    Loop
    rst.Close
    MsgBox "Atualização concluída com sucesso!", vbInformation, "Processamento Concluído!"
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    
        'Indicando ao usuário a conexão
        If InternetAutodial(INTERNET_AUTODIAL_FORCE_ONLINE, 0) Then
            MsgBox "Você esta conectado!", vbInformation
        'inicia a discagem automaticamente
        ElseIf InternetAutodial(INTERNET_AUTODIAL_FORCE_UNATTENDED, 0) Then
            MsgBox "Você já esta conectado!", vbInformation
        End If
        
    
    
    
    
    'If ValidaCampos Then
    '    AtivaBotoes
    '    If lOpcao = 1 Then
    '        gSQL = "INSERT INTO Bomba ( Empresa, Codigo, [Tipo de Combustivel], [Preco de Custo], [Preco de Venda], [Numero do Tanque], [Numero da Ilha], [Tipo de Preco], [Codigo do Produto] ) VALUES ( "
    '        Call sqlNumero(1, CStr(g_empresa), ", ")
    '        Call sqlNumero(1, txt_codigo, ", ")
    '        Call sqlTexto(1, dtcbo_combustivel.BoundText, ", ")
    '        Call sqlValor4(1, msk_custo, ", ")
    '        Call sqlValor4(1, msk_venda, ", ")
    '        Call sqlTexto(1, txt_tanque, ", ")
    '        Call sqlNumero(1, txt_numero_ilha, ", ")
    '        If cbo_tipo_preco.ListIndex = 0 Then
    '            Call sqlTexto(1, "V", ", ")
    '        ElseIf cbo_tipo_preco.ListIndex = 1 Then
    '            Call sqlTexto(1, "P", ", ")
    '        End If
    '        Call sqlNumero(1, dtcbo_produto.BoundText, " )")
    '        cnnSGP.Execute gSQL
    '    ElseIf lOpcao = 2 Then
    '        gSQL = "UPDATE Bomba SET "
    '        Call sqlTexto(2, "[Tipo de Combustivel] = ", dtcbo_combustivel.BoundText)
    '        Call sqlValor4(2, ", [Preco de Custo] = ", msk_custo)
    '        Call sqlValor4(2, ", [Preco de Venda] = ", msk_venda)
    '        Call sqlTexto(2, ", [Numero do Tanque] = ", txt_tanque)
    '        Call sqlNumero(2, ", [Numero da Ilha] = ", txt_numero_ilha)
    '        If cbo_tipo_preco.ListIndex = 0 Then
    '            Call sqlTexto(2, ", [Tipo de Preco] = ", "V")
    '        ElseIf cbo_tipo_preco.ListIndex = 1 Then
    '            Call sqlTexto(2, ", [Tipo de Preco] = ", "P")
    '        End If
    '        Call sqlNumero(2, ", [Codigo do Produto] = ", dtcbo_produto.BoundText)
    '        gSQL = gSQL & " WHERE Empresa = " & g_empresa
    '        gSQL = gSQL & " AND Codigo = " & lCodigo
    '        cnnSGP.Execute gSQL
    '    End If
    '    lCodigo = txt_codigo
    '    lOpcao = 0
    '    Call BuscaRegistro(lCodigo)
    '    cmd_novo.SetFocus
    'End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_bomba.Name, "Bombaa"
    Exit Sub
End Sub
Function PreparaSQL(xCondicao As String, xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Codigo, [Tipo de Combustivel], [Preco de Custo], [Preco de Venda], [Numero do Tanque], [Numero da Ilha], [Tipo de Preco], [Codigo do Produto] FROM Bomba WHERE Empresa = " & g_empresa & " " & xCondicao & " " & xTipoOrdem)
End Function
Function ValidaCampos() As Integer
    'ValidaCampos = False
    'If Not Val(txt_codigo) > 0 Then
    '    MsgBox "Informe o codigo do Bico.", 64, "Atenção!"
    '    txt_codigo.SetFocus
    'ElseIf IsNull(dtcbo_combustivel.SelectedItem) Then
    '    MsgBox "Escolha o Combustível.", 64, "Atenção!"
    '    dtcbo_combustivel.SetFocus
    'ElseIf Not fValidaValor4(msk_custo) > 0 Then
    '    MsgBox "Informe o Preço de Custo.", 64, "Atenção!"
    '    msk_custo.SetFocus
    'ElseIf Not fValidaValor4(msk_venda) > 0 Then
    '    MsgBox "Informe o Preço de Venda.", 64, "Atenção!"
    '    msk_venda.SetFocus
    'ElseIf txt_tanque = "" Then
    '    MsgBox "Informe o Número do Tanque.", 64, "Atenção!"
    '    txt_tanque.SetFocus
    'ElseIf Not Val(txt_numero_ilha) > 0 Then
    '    MsgBox "O número da ilha deve ser maior que 0.", 64, "Atenção!"
    '    txt_numero_ilha.SetFocus
    'ElseIf cbo_tipo_preco.ListIndex = -1 Then
    '    MsgBox "Escolha um tipo de preço.", 64, "Atenção!"
    '    cbo_tipo_preco.SetFocus
    'ElseIf IsNull(dtcbo_produto.SelectedItem) Then
    '    MsgBox "Escolha um produto.", 64, "Atenção!"
    '    dtcbo_produto.SetFocus
    'Else
    '    ValidaCampos = True
    'End If
End Function
Private Sub cmd_sair_Click()
    'desconecta
    If InternetAutodialHangup(0) Then
        MsgBox "Você esta desconectado", vbInformation
    End If
    Unload Me
End Sub
Private Sub Form_Activate()
    'If lFlagCadastroBomba = 0 Then
    '    DesativaBotoes
    '    If BuscaDados Then
    '        AtivaBotoes
    '    Else
    '        cmd_novo.Enabled = True
    '        cmd_sair.Enabled = True
    '    End If
    '    cmd_novo.SetFocus
    'Else
    '    lFlagCadastroBomba = 0
    'End If
End Sub
Private Sub Form_Deactivate()
    'lFlagCadastroBomba = 1
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    lTabela = "Contas_Pagar"
    txt_destinatario.Text = "tasso_cerrado@uol.com.br"
    txt_assunto.Text = "Contas a Pagar"
    txt_mensagem.Text = UCase(g_nome_empresa) & "|@|" & lTabela
    txt_anexo.Text = gDrive & gDiretorioData & lTabela & ".XML"
    
    'Set tbl_bomba = bd_sgp.OpenTable("Bomba")
    'Set tbl_combustivel = bd_sgp.OpenTable("Combustivel")
    'tbl_bomba.Index = "id_codigo"
    'tbl_combustivel.Index = "id_codigo"
    'dta_combustivel.RecordSource = "Select * From Combustivel Where Empresa = " & g_empresa & " Order By Nome"
    'dta_combustivel.Refresh
    'adodc_combustivel.ConnectionString = gConnectionString
    'adodc_combustivel.RecordSource = "SELECT Codigo, Nome FROM Combustivel ORDER BY Nome"
    'adodc_combustivel.Refresh
    'adodc_produto.ConnectionString = gConnectionString
    'adodc_produto.RecordSource = "SELECT Codigo, Nome FROM Produto WHERE Inativo = FALSE AND [Codigo do Grupo] = 4 ORDER BY Nome"
    'adodc_produto.Refresh
    'PreencheCboTipoPreco
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
