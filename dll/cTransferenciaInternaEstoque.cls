VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTransfInternaEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mData As Date
Private mPeriodo As Integer
Private mNumeroIlha As Integer
Private mCodigoSubEstoqueEntrada As Integer
Private mCodigoProduto As Long
Private mCodigoFuncionario As Integer
Private mCodigoSubEstoqueSaida As Integer
Private mQuantidade As Currency
Private mTransferido As Boolean

Private rsTransferenciaInternaEstoque As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get Data() As Date
    Data = mData
End Property
Public Property Let Data(ByVal Valor As Date)
    mData = Valor
End Property


Public Property Get Periodo() As Integer
    Periodo = mPeriodo
End Property
Public Property Let Periodo(ByVal Valor As Integer)
    mPeriodo = Valor
End Property


Public Property Get NumeroIlha() As Integer
    NumeroIlha = mNumeroIlha
End Property
Public Property Let NumeroIlha(ByVal Valor As Integer)
    mNumeroIlha = Valor
End Property


Public Property Get CodigoSubEstoqueEntrada() As Integer
    CodigoSubEstoqueEntrada = mCodigoSubEstoqueEntrada
End Property
Public Property Let CodigoSubEstoqueEntrada(ByVal Valor As Integer)
    mCodigoSubEstoqueEntrada = Valor
End Property


Public Property Get CodigoProduto() As Long
    CodigoProduto = mCodigoProduto
End Property
Public Property Let CodigoProduto(ByVal Valor As Long)
    mCodigoProduto = Valor
End Property


Public Property Get CodigoFuncionario() As Integer
    CodigoFuncionario = mCodigoFuncionario
End Property
Public Property Let CodigoFuncionario(ByVal Valor As Integer)
    mCodigoFuncionario = Valor
End Property


Public Property Get CodigoSubEstoqueSaida() As Integer
    CodigoSubEstoqueSaida = mCodigoSubEstoqueSaida
End Property
Public Property Let CodigoSubEstoqueSaida(ByVal Valor As Integer)
    mCodigoSubEstoqueSaida = Valor
End Property


Public Property Get Quantidade() As Currency
    Quantidade = mQuantidade
End Property
Public Property Let Quantidade(ByVal Valor As Currency)
    mQuantidade = Valor
End Property


Public Property Get Transferido() As Boolean
    Transferido = mTransferido
End Property
Public Property Let Transferido(ByVal Valor As Boolean)
    mTransferido = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pNumeroIlha As Integer, ByVal pCodigoSubEstoqueEntrada As Integer, ByVal pCodigoProduto As Long, ByVal pCodigoFuncionario As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND Periodo = " & pPeriodo
    xCondicao = xCondicao & " AND [Numero da Ilha] = " & pNumeroIlha
    xCondicao = xCondicao & " AND [Codigo do SubEstoque de Entrada] = " & pCodigoSubEstoqueEntrada
    xCondicao = xCondicao & " AND [Codigo do Produto] = " & pCodigoProduto
    xCondicao = xCondicao & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsTransferenciaInternaEstoque.Close
    Set rsTransferenciaInternaEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarAnterior = False
    xRegIgual = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data <= " & preparaData(mData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Data DESC, Periodo DESC, [Numero da Ilha] DESC, [Codigo do SubEstoque de Entrada] DESC, [Codigo do Produto] DESC, [Codigo do Funcionario] DESC")
    If Localizar(100) Then
        If rsTransferenciaInternaEstoque.RecordCount > 1 Then
            Do Until rsTransferenciaInternaEstoque.EOF
                If xRegIgual = False Then
                    If rsTransferenciaInternaEstoque("Data").Value = mData And rsTransferenciaInternaEstoque("Periodo").Value = mPeriodo And rsTransferenciaInternaEstoque("Numero da Ilha").Value = mNumeroIlha And rsTransferenciaInternaEstoque("Codigo do SubEstoque de Entrada").Value = mCodigoSubEstoqueEntrada And rsTransferenciaInternaEstoque("Codigo do Produto").Value = mCodigoProduto And rsTransferenciaInternaEstoque("Codigo do Funcionario").Value = mCodigoFuncionario Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsTransferenciaInternaEstoque.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsTransferenciaInternaEstoque.Close
    Set rsTransferenciaInternaEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal pNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL(" WHERE Nome = " & preparaTexto(pNome), "")
    LocalizarNome = Localizar(1)
    If LocalizarNome Then
        AtribuiValor
    End If
    rsTransferenciaInternaEstoque.Close
    Set rsTransferenciaInternaEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, Periodo ASC, [Numero da Ilha] ASC, [Codigo do SubEstoque de Entrada] ASC, [Codigo do Produto] ASC, [Codigo do Funcionario] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsTransferenciaInternaEstoque.Close
    Set rsTransferenciaInternaEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarProximo = False
    xRegIgual = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data >= " & preparaData(mData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, Periodo ASC, [Numero da Ilha] ASC, [Codigo do SubEstoque de Entrada] ASC, [Codigo do Produto] ASC, [Codigo do Funcionario] ASC")
    If Localizar(100) Then
        If rsTransferenciaInternaEstoque.RecordCount > 1 Then
            Do Until rsTransferenciaInternaEstoque.EOF
                If xRegIgual = False Then
                    If rsTransferenciaInternaEstoque("Data").Value = mData And rsTransferenciaInternaEstoque("Periodo").Value = mPeriodo And rsTransferenciaInternaEstoque("Numero da Ilha").Value = mNumeroIlha And rsTransferenciaInternaEstoque("Codigo do SubEstoque de Entrada").Value = mCodigoSubEstoqueEntrada And rsTransferenciaInternaEstoque("Codigo do Produto").Value = mCodigoProduto And rsTransferenciaInternaEstoque("Codigo do Funcionario").Value = mCodigoFuncionario Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsTransferenciaInternaEstoque.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsTransferenciaInternaEstoque.Close
    Set rsTransferenciaInternaEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, Periodo DESC, [Numero da Ilha] DESC, [Codigo do SubEstoque de Entrada] DESC, [Codigo do Produto] DESC, [Codigo do Funcionario] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsTransferenciaInternaEstoque.Close
    Set rsTransferenciaInternaEstoque = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO TransferenciaInternaEstoque ( Empresa, Data, Periodo, [Numero da Ilha], [Codigo do SubEstoque de Entrada], "
    gSQL = gSQL & "[Codigo do Produto], [Codigo do Funcionario], [Codigo do SubEstoque de Saida], Quantidade, Transferido ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlNumero(1, mPeriodo, ", ")
    Call sqlNumero(1, mNumeroIlha, ", ")
    Call sqlNumero(1, mCodigoSubEstoqueEntrada, ", ")
    Call sqlNumero(1, mCodigoProduto, ", ")
    Call sqlNumero(1, mCodigoFuncionario, ", ")
    Call sqlNumero(1, mCodigoSubEstoqueSaida, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlBoolean(1, mTransferido, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pNumeroIlha As Integer, ByVal pCodigoSubEstoqueEntrada As Integer, ByVal pCodigoProduto As Long, ByVal pCodigoFuncionario As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE TransferenciaInternaEstoque SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlNumero(2, ", Periodo = ", mPeriodo)
    Call sqlNumero(2, ", [Numero da Ilha] = ", mNumeroIlha)
    Call sqlNumero(2, ", [Codigo do SubEstoque de Entrada] = ", mCodigoSubEstoqueEntrada)
    Call sqlNumero(2, ", [Codigo do Produto] = ", mCodigoProduto)
    Call sqlNumero(2, ", [Codigo do Funcionario] = ", mCodigoFuncionario)
    Call sqlNumero(2, ", [Codigo do SubEstoque de Saida] = ", mCodigoSubEstoqueSaida)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlBoolean(2, ", Transferido = ", mTransferido)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Periodo = " & pPeriodo
    gSQL = gSQL & " AND [Numero da Ilha] = " & pNumeroIlha
    gSQL = gSQL & " AND [Codigo do SubEstoque de Entrada] = " & pCodigoSubEstoqueEntrada
    gSQL = gSQL & " AND [Codigo do Produto] = " & pCodigoProduto
    gSQL = gSQL & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pNumeroIlha As Integer, ByVal pCodigoSubEstoqueEntrada As Integer, ByVal pCodigoProduto As Long, ByVal pCodigoFuncionario As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM TransferenciaInternaEstoque"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Periodo = " & pPeriodo
    gSQL = gSQL & " AND [Numero da Ilha] = " & pNumeroIlha
    gSQL = gSQL & " AND [Codigo do SubEstoque de Entrada] = " & pCodigoSubEstoqueEntrada
    gSQL = gSQL & " AND [Codigo do Produto] = " & pCodigoProduto
    gSQL = gSQL & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, Periodo, [Numero da Ilha], [Codigo do SubEstoque de Entrada], [Codigo do Produto], [Codigo do Funcionario], [Codigo do SubEstoque de Saida], Quantidade, Transferido FROM TransferenciaInternaEstoque" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsTransferenciaInternaEstoque = New adodb.Recordset
    rsTransferenciaInternaEstoque.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsTransferenciaInternaEstoque.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsTransferenciaInternaEstoque.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsTransferenciaInternaEstoque("Empresa").Value
    Data = rsTransferenciaInternaEstoque("Data").Value
    Periodo = rsTransferenciaInternaEstoque("Periodo").Value
    NumeroIlha = rsTransferenciaInternaEstoque("Numero da Ilha").Value
    CodigoSubEstoqueEntrada = rsTransferenciaInternaEstoque("Codigo do SubEstoque de Entrada").Value
    CodigoProduto = rsTransferenciaInternaEstoque("Codigo do Produto").Value
    CodigoFuncionario = rsTransferenciaInternaEstoque("Codigo do Funcionario").Value
    CodigoSubEstoqueSaida = rsTransferenciaInternaEstoque("Codigo do SubEstoque de Saida").Value
    Quantidade = rsTransferenciaInternaEstoque("Quantidade").Value
    Transferido = rsTransferenciaInternaEstoque("Transferido").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

