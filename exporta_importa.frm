VERSION 5.00
Object = "{00028C21-0000-0000-0000-000000000046}#4.0#0"; "TDBG32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form exporta_importa 
   Caption         =   "Exporta / Importa Dados do/para Sistema Gerenciador de Posto Cerrado"
   ClientHeight    =   7800
   ClientLeft      =   2310
   ClientTop       =   885
   ClientWidth     =   8820
   Icon            =   "exporta_importa.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "exporta_importa.frx":0442
   ScaleHeight     =   7800
   ScaleWidth      =   8820
   Begin VB.ComboBox cboUnidadeSecundaria 
      Height          =   315
      Left            =   3180
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Frame FrmSelecionar 
      Caption         =   "S&elecionar por..."
      Height          =   1155
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   6855
      Begin VB.ComboBox cbo_campo 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   5775
      End
      Begin VB.TextBox txt_condicao 
         Height          =   285
         Left            =   3120
         TabIndex        =   9
         Top             =   720
         Width           =   3615
      End
      Begin VB.ComboBox cbo_operador 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "&Campo"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "Co&ndição"
         Height          =   255
         Left            =   2340
         TabIndex        =   8
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "O&perador"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   780
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7020
      Top             =   1140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Origem para Importar"
      FileName        =   "*.mdb"
      InitDir         =   "a:\"
   End
   Begin VB.CommandButton cmd_importa_dados 
      Caption         =   "&Importa Dados"
      Height          =   375
      Left            =   7260
      Picture         =   "exporta_importa.frx":0488
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Importa dados do disquete (A:)."
      Top             =   1440
      Width           =   1515
   End
   Begin VB.CommandButton cmd_exporta_dados 
      Caption         =   "Exporta &Dados"
      Height          =   375
      Left            =   7260
      Picture         =   "exporta_importa.frx":187A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Exporta dados para disquete (A:)."
      Top             =   1020
      Width           =   1515
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   7020
      Picture         =   "exporta_importa.frx":2C6C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Inicia a pesquisa selecionada."
      Top             =   120
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   7980
      Picture         =   "exporta_importa.frx":3F46
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   120
      Width           =   795
   End
   Begin VB.Data dta_tabela 
      Connect         =   "Access"
      DatabaseName    =   "Sgp_data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   7500
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame frmTabela 
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.ComboBox cbo_tabela 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   180
         Width           =   5775
      End
      Begin VB.Label Label1 
         Caption         =   "&Tabelas"
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   675
      End
   End
   Begin TrueDBGrid.TDBGrid TDBGrid1 
      Bindings        =   "exporta_importa.frx":5220
      Height          =   5415
      Left            =   60
      OleObjectBlob   =   "exporta_importa.frx":5239
      TabIndex        =   10
      Top             =   2400
      Width           =   8715
   End
   Begin VB.Label Label5 
      Caption         =   "&Unidade secundário"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1980
      Width           =   3015
   End
End
Attribute VB_Name = "exporta_importa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lIndicesTmp As Object
Dim snp_tabelas As Snapshot
Dim snp_campos As Snapshot
Dim l_nome_tabela As Table
Dim tbl_tabela As Table
Dim tbl_tabela_disco As Table
Dim mfldCurrFld As Field
Dim mindCurrInd As Index
Dim lNomeBancoNovo As String
Dim l_campo As String
Dim l_arquivo As String
Dim l_condicao As String
Dim l_ordem As String
Dim lSQL As String
    
Dim lNomeArquivo As String
Dim lNomeIndex As String
Dim lIndexUnico As Boolean
Dim lComposicaoIndex(0 To 15) As String
Dim lRstDataIndex As adodb.Recordset
Dim cat As ADOX.Catalog
Dim CatNovo As ADOX.Catalog
Dim tbl As New ADOX.Table
Dim tblNova As New ADOX.Table
Dim Col As New ADOX.Column
Dim ColNova As New ADOX.Column
Dim Idx As New ADOX.Index
Dim IdxNovo As New ADOX.Index
Dim rstNovo As New adodb.Recordset
Dim rst As New adodb.Recordset
Dim fldNovo As adodb.Field
Dim cnnNovo As adodb.Connection


Private Sub AtualizaGrid()
    On Error GoTo ErroConsulta
    Dim x_operando As String
    Dim x_condicao As String
    Dim x_data As Date
    x_condicao = txt_condicao
    snp_campos.MoveFirst
    Do Until snp_campos.EOF
        If Mid(cbo_campo.Text, 1, 1) = "[" And snp_campos!name = Mid(cbo_campo.Text, 2, Len(cbo_campo.Text) - 2) Or snp_campos!name = cbo_campo.Text Then
            If snp_campos!Type = 8 Then
                x_condicao = "#" & Format(CDate(x_condicao), "yyyy/mm/dd") & "#"
            ElseIf snp_campos!Type = 10 And cbo_operador.Text = "Semelhante" Then
                x_condicao = Chr(34) & "*" & x_condicao & "*" & Chr(34)
            ElseIf snp_campos!Type = 10 And cbo_operador.Text <> "Semelhante" Then
                x_condicao = Chr(34) & x_condicao & Chr(34)
            End If
            Exit Do
        End If
        snp_campos.MoveNext
    Loop
    If cbo_operador.Text = "Diferente" Then
        x_operando = "<>"
    ElseIf cbo_operador.Text = "Igual" Then
        x_operando = "="
    ElseIf cbo_operador.Text = "Maior" Then
        x_operando = ">"
    ElseIf cbo_operador.Text = "Maior Igual" Then
        x_operando = ">="
    ElseIf cbo_operador.Text = "Menor" Then
        x_operando = "<"
    ElseIf cbo_operador.Text = "Menor Igual" Then
        x_operando = "<="
    ElseIf cbo_operador.Text = "Semelhante" Then
        x_operando = "Like"
    End If
    If ValidaCampos Then
        l_campo = "Select " & "* "
        l_arquivo = "From " & cbo_tabela.Text & " "
        If cbo_campo.ListIndex <> -1 Then
            l_condicao = "Where "
            l_condicao = l_condicao & cbo_tabela.Text & "." & cbo_campo.Text & " " & x_operando & " " & x_condicao
        Else
            l_condicao = ""
        End If
        'l_ordem = "order by cheques.emitente"
        l_ordem = ""
        lSQL = l_campo & l_arquivo & l_condicao & l_ordem
'TESTE        Set l_nome_tabela = bd_sgp.CreateSnapshot(lSQL)
        dta_tabela.RecordSource = lSQL
        dta_tabela.Refresh
    End If
    Exit Sub
ErroConsulta:
    If Err = 3075 Then
        MsgBox "Condição inválida.", vbExclamation, "Erro de Consulta"
        Exit Sub
    End If
    Exit Sub
End Sub
Private Function Condicao() As String
    Condicao = txt_condicao.Text
    snp_campos.MoveFirst
    Do Until snp_campos.EOF
        If snp_campos!name = cbo_campo.Text Then
            If snp_campos!Type = 8 Then
                Condicao = "#" & Format(CDate(txt_condicao.Text), "yyyy/mm/dd") & "#"
            ElseIf (snp_campos!Type = 10 Or snp_campos!Type = 202) And cbo_operador.Text = "Semelhante" Then
                Condicao = Chr(39) & "%" & txt_condicao.Text & "%" & Chr(39)
            ElseIf (snp_campos!Type = 10 Or snp_campos!Type = 200 Or snp_campos!Type = 202) And cbo_operador.Text <> "Semelhante" Then
                Condicao = Chr(39) & txt_condicao.Text & Chr(39)
            End If
            Exit Do
        End If
        snp_campos.MoveNext
    Loop
End Function
Private Function CriaBancoDados() As Boolean
    On Error GoTo RotinaDeErro
    
    CriaBancoDados = False
    Set CatNovo = New ADOX.Catalog
    CatNovo.Create "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=4;Data Source=" & lNomeArquivo
    CriaBancoDados = True
    Exit Function
    
RotinaDeErro:
    MsgBox "Erro ao Criar o Banco de Dados: " & lNomeArquivo, vbInformation, "CriaBancoDados"
    Exit Function
End Function
Private Function CriaCampos() As Boolean
    On Error GoTo RotinaDeErro
    Dim xRequired As Boolean
    Dim rsInfo As adodb.Recordset
    
    CriaCampos = False
    For Each Col In tbl.Columns
        Set rsInfo = cnnSGP.OpenSchema(adSchemaColumns, Array(Empty, Empty, lNomeBancoNovo, Col.name))
        xRequired = False
        ''If Not rsInfo.EOF Then
        ''    MsgBox rsInfo!COLUMN_NAME & " - " & rsInfo!IS_NULLABLE
        ''    xRequired = rsInfo!IS_NULLABLE
        ''End If
        'Cria Campos (Colunas)
        ''If Col.Type = adDBDate Or Col.Type = adDate Or Col.Type = adDBTime Or Col.Type = adDBTimeStamp Then
        ''    TblNova.Columns.Append Col.Name, adDate
        ''ElseIf Col.Type = adVarChar Or Col.Type = adVarWChar Then
        ''    TblNova.Columns.Append Col.Name, adVarWChar, Col.DefinedSize
        ''ElseIf Col.Type = adSmallInt Then
        ''    TblNova.Columns.Append Col.Name, adSmallInt, Col.DefinedSize
        ''Else
        ''    TblNova.Columns.Append Col.Name, Col.Type
        ''End If
        ''MsgBox IIf((Col.Attributes And adColNullable) = adColNullable, "Sim", "Não") & " "
        ''TblNova.Columns(Col.Name).Attributes = Col.Attributes
        'TblNova.Columns(Col.Name).Properties("Jet OLEDB:Allow Zero Length").Value = True
        'Set novaTabela = catDB.Tables(xNomeTabela)
        'novaTabela.Columns("Observacao 1").Properties("Jet OLEDB:Allow Zero Length").Value = True
        'novaTabela.Columns("Observacao 2").Properties("Jet OLEDB:Allow Zero Length").Value = True
        'Col.Properties("Nullable") = True
        'Col.Properties("Jet OLEDB:Allow Zero Length") = True
        'novaTabela.Columns("Observacao 1").Properties("Jet OLEDB:Allow Zero Length").Value = True
        'MsgBox Col.Name & " - " & Col("").Properties("Jet OLEDB:Allow Zero Length").Value
        Set ColNova = New ADOX.Column
        ColNova.name = Col.name
        ColNova.Type = Col.Type
        If Col.Type = adDBTimeStamp Then
            ColNova.Type = adDate
        ElseIf Col.Type = adNumeric Then
            ColNova.Type = adCurrency
        End If
        Set ColNova.ParentCatalog = CatNovo
        ColNova.Attributes = Col.Attributes
        ColNova.DefinedSize = Col.DefinedSize
        ColNova.NumericScale = Col.NumericScale
        ColNova.Precision = Col.Precision
        tblNova.Columns.Append ColNova
    Next
    Set rsInfo = Nothing
    CriaCampos = True
    Exit Function
    
RotinaDeErro:
    MsgBox "Erro ao criar campos da tabela!", vbInformation, "CriaCampos"
    Exit Function
End Function
Private Sub CriaIndex()
Dim i As Integer
    If tbl.Indexes.Count >= 2 Then
        For Each Idx In tbl.Indexes
            If Idx.name = lNomeIndex Then
                'definindo as propriedades do índice
                Set IdxNovo = New ADOX.Index
                'IdxNovo.Clustered = Idx.Clustered
                'IdxNovo.IndexNulls = False
                IdxNovo.name = Idx.name
                IdxNovo.PrimaryKey = False 'Idx.PrimaryKey
                IdxNovo.Unique = Idx.Unique
                'IdxNovo.Name = lNomeIndex
                'IdxNovo.PrimaryKey = True
                'IdxNovo.Unique = lIndexUnico
                For i = 0 To 15
                    If lComposicaoIndex(i) <> "" Then
                        IdxNovo.Columns.Append lComposicaoIndex(i)
                    Else
                        Exit For
                    End If
                Next
                'Criando Índice
                tblNova.Indexes.Append IdxNovo
                Set IdxNovo = Nothing
                Exit For
            End If
        Next
    End If
End Sub
Private Function CriaTabela() As Boolean
    On Error GoTo RotinaDeErro
    
    CriaTabela = False
    Set cat = New ADOX.Catalog
    Set cat.ActiveConnection = cnnSGP
    For Each tbl In cat.Tables
        If tbl.name = lNomeBancoNovo Then
            'Cria Nova Tabela
            Set tblNova = New ADOX.Table
            tblNova.name = lNomeBancoNovo
            CatNovo.Tables.Append tblNova
            Exit For
        End If
    Next
    CriaTabela = True
    Exit Function
    
RotinaDeErro:
    MsgBox "Erro ao Criar a tabela: " & lNomeBancoNovo, vbInformation, "CriaTabela"
    Exit Function
End Function
Private Sub ExportaDados()
    Dim xSQL As String
    On Error GoTo RotinaDeErro
    
    xSQL = ""
    xSQL = xSQL & "SELECT *"
    xSQL = xSQL & "  FROM " & lNomeBancoNovo
    xSQL = xSQL & " WHERE " & cbo_campo.Text
    xSQL = xSQL & " " & Operando & " " & Condicao
    
    Set cnnNovo = New adodb.Connection
    cnnNovo.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & lNomeArquivo
    For Each tbl In cat.Tables
        If tbl.name = lNomeBancoNovo Then
            rstNovo.Open tbl.name, cnnNovo, adOpenKeyset, adLockOptimistic, adCmdTableDirect
            'rst.Open tblNova.Name, cnnSGP, adOpenKeyset, adLockOptimistic, adCmdTableDirect
            rst.Open xSQL, cnnSGP, adOpenForwardOnly, adLockReadOnly
            If Not rst.BOF And Not rst.EOF Then
                rst.MoveFirst
                Do Until rst.EOF
                    rstNovo.AddNew
                    For Each fldNovo In rstNovo.Fields
                        rstNovo.Fields(fldNovo.name).Value = rst.Fields(fldNovo.name).Value
                    Next
                    rstNovo.Update
                    rst.MoveNext
                Loop
            End If
            rst.Close
            rstNovo.Close
            cnnNovo.Close
            Set rst = Nothing
            Set rstNovo = Nothing
            Set cnnNovo = Nothing
            'MsgBox "Dados exportado com sucesso!", vbInformation, "Operação Concluída"
            Exit For
        End If
    Next
    Exit Sub
    
RotinaDeErro:
    rst.Close
    rstNovo.Close
    cnnNovo.Close
    Set rst = Nothing
    Set rstNovo = Nothing
    Set cnnNovo = Nothing
    MsgBox "Erro ao gravar dados!", vbInformation, "ExportaDados"
    Exit Sub
End Sub
Private Sub Finaliza()
    Set cat = Nothing
End Sub
Private Function GravaRst() As Boolean
    On Error GoTo RotinaDeErro
    rst.Update
    GravaRst = True
    Exit Function
    
RotinaDeErro:
    GravaRst = False
    rst.CancelUpdate
End Function
Private Sub ImportaDados()
    Dim xSQL As String
    On Error GoTo RotinaDeErro
    
    lNomeBancoNovo = cbo_tabela.Text
    lNomeArquivo = cboUnidadeSecundaria.Text & "\" & lNomeBancoNovo & ".MDB"
    If Not gArqTxt.FileExists(lNomeArquivo) Then
        MsgBox "O banco de dados " & lNomeArquivo & " não existe.", vbInformation, "Origem Inexistente!"
        Exit Sub
    End If
    
    If (MsgBox("Deseja substituir os dados já existentes?", vbYesNo + vbQuestion + vbDefaultButton2, "Substituição de Dados!")) = vbYes Then
        xSQL = ""
        xSQL = xSQL & "DELETE "
        xSQL = xSQL & "  FROM " & lNomeBancoNovo
        xSQL = xSQL & " WHERE " & cbo_campo.Text
        xSQL = xSQL & " " & Operando & " " & Condicao
        cnnSGP.Execute xSQL
    End If
    
    xSQL = ""
    xSQL = xSQL & "SELECT *"
    xSQL = xSQL & "  FROM " & lNomeBancoNovo
    xSQL = xSQL & " WHERE " & cbo_campo.Text
    xSQL = xSQL & " " & Operando & " " & Condicao
    
    Set cnnNovo = New adodb.Connection
    cnnNovo.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & lNomeArquivo
    For Each tbl In cat.Tables
        If tbl.name = lNomeBancoNovo Then
            rst.Open tbl.name, cnnSGP, adOpenKeyset, adLockOptimistic, adCmdTableDirect
            rstNovo.Open xSQL, cnnNovo, adOpenForwardOnly, adLockReadOnly
            If Not rstNovo.BOF And Not rstNovo.EOF Then
                rstNovo.MoveFirst
                Do Until rstNovo.EOF
                    rst.AddNew
                    For Each fldNovo In rst.Fields
                        rst.Fields(fldNovo.name).Value = rstNovo.Fields(fldNovo.name).Value
                    Next
                    If Not GravaRst Then
                        MsgBox "Não foi possível gravar este registro.", vbInformation, "Registro Existente!"
                    End If
                    rstNovo.MoveNext
                Loop
            End If
            rst.Close
            rstNovo.Close
            cnnNovo.Close
            Set rst = Nothing
            Set rstNovo = Nothing
            Set cnnNovo = Nothing
            MsgBox "Dados importado com sucesso!", vbInformation, "Operação Concluída"
            Exit For
        End If
    Next
    Exit Sub
    
RotinaDeErro:
    rst.Close
    rstNovo.Close
    cnnNovo.Close
    Set rst = Nothing
    Set rstNovo = Nothing
    Set cnnNovo = Nothing
    MsgBox "Erro ao gravar dados!", vbInformation, "ExportaDados"
    Exit Sub
End Sub
Private Function MoveBdParaDestino() As Boolean
    Dim xArquivoDestino As String
    On Error GoTo RotinaDeErro
    
    MoveBdParaDestino = False
    xArquivoDestino = lNomeArquivo
    Mid(xArquivoDestino, 1, 2) = cboUnidadeSecundaria.Text
    
    
    'Testa se a unidade de destino é a mesma da origem
    If gDrive = cboUnidadeSecundaria.Text Then
        MsgBox "Para que esta operação seja efetuada com sucesso," & Chr(10) & "este programa será finalizado automaticamente.", vbInformation, "Atenção!"
        Call WriteINI("ARQUIVOS", "Origem", lNomeArquivo, "c:\transfere.ini")
        Call WriteINI("ARQUIVOS", "Destino", xArquivoDestino, "c:\transfere.ini")
        End
    End If

    'Deleta o Arquivo de Destino Caso Exista
    If gArqTxt.FileExists(xArquivoDestino) Then
        Call gArqTxt.DeleteFile(xArquivoDestino)
    End If
    
    'Copia o Arquivo de Origem para o Destino
    Call gArqTxt.CopyFile(lNomeArquivo, xArquivoDestino)

    
    'Copia um arquivo encima do Arquivo de Origem
    If gArqTxt.FileExists(gArquivoIni) Then
        Call gArqTxt.CopyFile(gArquivoIni, lNomeArquivo, True)
    End If
    
    'Deleta o Arquivo de Origem
    Call gArqTxt.DeleteFile(lNomeArquivo)

    MoveBdParaDestino = True
    Exit Function
    
RotinaDeErro:

End Function
Private Function Operando() As String
    Operando = ""
    If cbo_operador.Text = "Diferente" Then
        Operando = "<>"
    ElseIf cbo_operador.Text = "Igual" Then
        Operando = "="
    ElseIf cbo_operador.Text = "Maior" Then
        Operando = ">"
    ElseIf cbo_operador.Text = "Maior Igual" Then
        Operando = ">="
    ElseIf cbo_operador.Text = "Menor" Then
        Operando = "<"
    ElseIf cbo_operador.Text = "Menor Igual" Then
        Operando = "<="
    ElseIf cbo_operador.Text = "Semelhante" Then
        Operando = "Like"
    End If
End Function
Private Sub PreencheCampos()
    Set snp_campos = l_nome_tabela.ListFields()
    'Set lIndices = l_nome_tabela.ListIndexes()
    Dim i As Integer
    Dim x_string As String
    cbo_campo.Clear
    snp_campos.MoveFirst
    Do Until snp_campos.EOF
        x_string = snp_campos!name
        For i = 1 To Len(snp_campos!name)
            If Mid(snp_campos!name, i, 1) = " " Then
                x_string = "[" & snp_campos!name & "]"
            End If
        Next
        cbo_campo.AddItem x_string
        snp_campos.MoveNext
    Loop
End Sub
Private Sub PreencheOperador()
    cbo_operador.Clear
    cbo_operador.AddItem "Diferente"
    cbo_operador.AddItem "Igual"
    cbo_operador.AddItem "Maior"
    cbo_operador.AddItem "Maior Igual"
    cbo_operador.AddItem "Menor"
    cbo_operador.AddItem "Menor Igual"
    cbo_operador.AddItem "Semelhante"
End Sub
Private Sub PreencheTabela()
    For Each tbl In cat.Tables
        If tbl.Type = "TABLE" Then
            cbo_tabela.AddItem tbl.name
        End If
    Next
End Sub
Private Sub PreencheUnidadeSecundaria()
    Dim DriveNum As Long
    cboUnidadeSecundaria.Clear
    For DriveNum = 0 To 25
        If CBool(GetLogicalDrives And (2 ^ DriveNum)) = True Then
            cboUnidadeSecundaria.AddItem Chr(Asc("A") + DriveNum) & ":"
        End If
    Next
    cboUnidadeSecundaria.ListIndex = 0
End Sub
Private Sub cbo_campo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_operador.SetFocus
    End If
End Sub
Private Sub cbo_operador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_condicao.SetFocus
    End If
End Sub
Private Sub cbo_tabela_Click()
    If cbo_tabela.ListIndex <> -1 Then
        Set l_nome_tabela = bd_sgp.OpenTable(cbo_tabela.Text)
        dta_tabela.RecordSource = cbo_tabela.Text
        TDBGrid1.Caption = cbo_tabela.Text
        PreencheCampos
    End If
End Sub
Private Sub cbo_tabela_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_campo.SetFocus
    End If
End Sub
Private Sub PreparaCriacaoIndice()
Dim i As Integer
Dim i2 As Integer
    Set lRstDataIndex = cnnSGP.OpenSchema(adSchemaIndexes, Array(Empty, Empty, Empty, Empty, lNomeBancoNovo))
    If Not lRstDataIndex.EOF Then
        With lRstDataIndex
            lNomeIndex = ""
            .MoveFirst
            Do While .EOF = False
                If .Fields(5).Value <> lNomeIndex Then
                    If lNomeIndex <> "" Then
                        CriaIndex
                    End If
                    For i = 0 To 15
                        lComposicaoIndex(i) = ""
                    Next
                    lNomeIndex = .Fields(5).Value
                    i = -1
                End If
                If .Fields(17).Value <> "" Then
                    'For i2 = 0 To 24
                    '    MsgBox i2 & " - " & .Fields(i2).Name & " - " & .Fields(i2).Value
                    'Next
                    lIndexUnico = .Fields(7).Value
                    i = i + 1
                    lComposicaoIndex(i) = .Fields(17).Value
                End If
                .MoveNext
            Loop
            CriaIndex
        End With
    End If
End Sub
Private Sub teste()
    Dim iCnt As Integer
    Dim sCurIndexName As String
    Dim sIndexFields As String
    Dim sTableName As String
    Dim cnnDataIX As adodb.Connection
    Dim rsDataIX As adodb.Recordset
    Dim i2 As Integer
    
    
    
    
    
    
    
    
    
    
    
    
    
    'Set cnnDataIX = New ADODB.Connection
    'cnnDataIX.ConnectionString = gConnectionString
    'cnnDataIX.Open
    sTableName = cbo_tabela.Text
    Set rsDataIX = cnnSGP.OpenSchema(adSchemaIndexes, Array(Empty, Empty, Empty, Empty, sTableName))
') ', Array(Empty, Empty, sTableName))
    
    
    
    
    If rsDataIX.BOF = True Or rsDataIX.EOF = True Then
        MsgBox "Warning:" & vbLf & "No Primary Key is assigned to Table" & vbLf & sTableName
    Else
        With rsDataIX
            .MoveFirst
            Do While .EOF = False
                For i2 = 0 To 24
                    MsgBox i2 & " - " & .Fields(i2).name & " - " & .Fields(i2).Value
                Next
                .MoveNext
            Loop
        End With
    End If
    rsDataIX.Close
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If cbo_tabela.ListIndex = -1 Then
        MsgBox "Selecione a tabela.", vbInformation, "Atenção!"
        cbo_tabela.SetFocus
    ElseIf cbo_campo.ListIndex = -1 Then
        MsgBox "Informe o campo a ser testado.", vbInformation, "Atenção!"
        cbo_campo.SetFocus
    ElseIf cbo_operador.ListIndex = -1 Then
        MsgBox "Informe o operando a ser testado.", vbInformation, "Atenção!"
        cbo_operador.SetFocus
    ElseIf txt_condicao = "" Then
        MsgBox "Informe a condição testada.", vbInformation, "Atenção!"
        txt_condicao.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_exporta_dados_Click()
    Dim i As Integer
    On Error GoTo FileError
    
    lNomeBancoNovo = cbo_tabela.Text
    Set tbl_tabela = bd_sgp.OpenTable(lNomeBancoNovo)
    lNomeArquivo = "C:\" & lNomeBancoNovo & ".MDB"
    
    'Deleta o Arquivo de Destino Caso Exista
    If gArqTxt.FileExists(lNomeArquivo) Then
        Call gArqTxt.DeleteFile(lNomeArquivo)
    End If
    
    If (MsgBox("Deseja Realmente Criar o Banco de Dados: " & lNomeArquivo, 4 + 32 + 256, "Criação de Banco de Dados!")) = 6 Then
        If CriaBancoDados Then
            If CriaTabela Then
                If CriaCampos Then
                    PreparaCriacaoIndice
                    Sleep 2000
                    ExportaDados
                    Set tblNova = Nothing
                    Set CatNovo = Nothing
                    Set lRstDataIndex = Nothing
                    If MoveBdParaDestino Then
                        MsgBox "O banco de dados foi criado com sucesso!", vbExclamation, "Banco de Dados Criado!"
                    End If
                    Exit Sub
                End If
            End If
        End If
        Set tblNova = Nothing
        Set CatNovo = Nothing
        Set lRstDataIndex = Nothing
    End If
    Exit Sub

FileError:
    MsgBox "Erro ao Criar o Banco de Dados: " & lNomeArquivo, vbInformation, "Operação sem sucesso!"
    Exit Sub
End Sub
Private Sub cmd_importa_dados_Click()
    ImportaDados
    Exit Sub
End Sub
Private Sub cmd_ok_Click()
    AtualizaGrid
    TDBGrid1.SetFocus
End Sub
Private Sub cmd_sair_Click()
    Finaliza
    Unload Me
End Sub
Private Sub Form_Activate()
    PreencheTabela
    PreencheOperador
    PreencheUnidadeSecundaria
    cbo_tabela.SetFocus
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    Set cat = New ADOX.Catalog
    Set CatNovo = New ADOX.Catalog
    Set cat.ActiveConnection = cnnSGP
End Sub
Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 98 Then
        KeyCode = 40
    ElseIf KeyCode = 99 Then
        KeyCode = 34
    ElseIf KeyCode = 104 Then
        KeyCode = 38
    ElseIf KeyCode = 105 Then
        KeyCode = 33
    End If
End Sub
Private Sub txt_condicao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok_Click
    End If
End Sub
