VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoCaixaPista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mData As Date
Private mNumeroMovimento As Integer
Private mValor As Currency
Private mNumeroDocumento As String
Private mCodigoHistorico As Integer
Private mComplemento As String
Private mNumeroContaDebito As String
Private mNumeroContaCredito As String
Private mTipoMovimento As Integer
Private mCodigoUsuario As Integer
Private mPeriodo As Integer
Private mNumeroIlha As Integer
Private mDadosInterno As String
Private mCodigoLancamentoPadrao As Integer
Private mDataDigitacao As Date
Private mHoraDigitacao As Date
Private mDataAlteracao As Date
Private mHoraAlteracao As Date

Private rsMovCaixaPista As New adodb.Recordset


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


Public Property Get NumeroMovimento() As Integer
    NumeroMovimento = mNumeroMovimento
End Property
Public Property Let NumeroMovimento(ByVal Valor As Integer)
    mNumeroMovimento = Valor
End Property


Public Property Get Valor() As Currency
    Valor = mValor
End Property
Public Property Let Valor(ByVal Valor As Currency)
    mValor = Valor
End Property


Public Property Get NumeroDocumento() As String
    NumeroDocumento = mNumeroDocumento
End Property
Public Property Let NumeroDocumento(ByVal Valor As String)
    mNumeroDocumento = Valor
End Property


Public Property Get CodigoHistorico() As Integer
    CodigoHistorico = mCodigoHistorico
End Property
Public Property Let CodigoHistorico(ByVal Valor As Integer)
    mCodigoHistorico = Valor
End Property


Public Property Get Complemento() As String
    Complemento = mComplemento
End Property
Public Property Let Complemento(ByVal Valor As String)
    mComplemento = Valor
End Property


Public Property Get NumeroContaDebito() As String
    NumeroContaDebito = mNumeroContaDebito
End Property
Public Property Let NumeroContaDebito(ByVal Valor As String)
    mNumeroContaDebito = Valor
End Property


Public Property Get NumeroContaCredito() As String
    NumeroContaCredito = mNumeroContaCredito
End Property
Public Property Let NumeroContaCredito(ByVal Valor As String)
    mNumeroContaCredito = Valor
End Property


Public Property Get TipoMovimento() As Integer
    TipoMovimento = mTipoMovimento
End Property
Public Property Let TipoMovimento(ByVal Valor As Integer)
    mTipoMovimento = Valor
End Property


Public Property Get CodigoUsuario() As Integer
    CodigoUsuario = mCodigoUsuario
End Property
Public Property Let CodigoUsuario(ByVal Valor As Integer)
    mCodigoUsuario = Valor
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


Public Property Get DadosInterno() As String
    DadosInterno = mDadosInterno
End Property
Public Property Let DadosInterno(ByVal Valor As String)
    mDadosInterno = Valor
End Property


Public Property Get CodigoLancamentoPadrao() As Integer
    CodigoLancamentoPadrao = mCodigoLancamentoPadrao
End Property
Public Property Let CodigoLancamentoPadrao(ByVal Valor As Integer)
    mCodigoLancamentoPadrao = Valor
End Property


Public Property Get DataDigitacao() As Date
    DataDigitacao = mDataDigitacao
End Property
Public Property Let DataDigitacao(ByVal Valor As Date)
    mDataDigitacao = Valor
End Property


Public Property Get HoraDigitacao() As Date
    HoraDigitacao = mHoraDigitacao
End Property
Public Property Let HoraDigitacao(ByVal Valor As Date)
    mHoraDigitacao = Valor
End Property


Public Property Get DataAlteracao() As Date
    DataAlteracao = mDataAlteracao
End Property
Public Property Let DataAlteracao(ByVal Valor As Date)
    mDataAlteracao = Valor
End Property


Public Property Get HoraAlteracao() As Date
    HoraAlteracao = mHoraAlteracao
End Property
Public Property Let HoraAlteracao(ByVal Valor As Date)
    mHoraAlteracao = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumeroMovimento As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND [Numero do Movimento] = " & pNumeroMovimento
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsMovCaixaPista.Close
    Set rsMovCaixaPista = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarRegistroEspecial(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pIlha As Integer, ByVal pComplemento As String, ByVal pNumeroConta As String, ByVal pDebitoCredito As String) As Boolean
    Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarRegistroEspecial = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    If pDebitoCredito = "D" Then
        xCondicao = xCondicao & " AND [Numero da Conta Debito] = " & preparaTexto(pNumeroConta)
    Else
        xCondicao = xCondicao & " AND [Numero da Conta Credito] = " & preparaTexto(pNumeroConta)
    End If
    xCondicao = xCondicao & " AND Periodo = " & pPeriodo
    xCondicao = xCondicao & " AND [Numero da Ilha] = " & pIlha
    xCondicao = xCondicao & " AND Complemento = " & preparaTexto(pComplemento)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Movimento] ASC")
    If Localizar(1) Then
        LocalizarRegistroEspecial = True
        AtribuiValor
    End If
    rsMovCaixaPista.Close
    Set rsMovCaixaPista = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarRegistroEspecialDoc(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pIlha As Integer, ByVal pComplemento As String, ByVal pNumeroDocumento As String, ByVal pNumeroConta As String, ByVal pDebitoCredito As String) As Boolean
    Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarRegistroEspecialDoc = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    If pDebitoCredito = "D" Then
        xCondicao = xCondicao & " AND [Numero da Conta Debito] = " & preparaTexto(pNumeroConta)
    Else
        xCondicao = xCondicao & " AND [Numero da Conta Credito] = " & preparaTexto(pNumeroConta)
    End If
    xCondicao = xCondicao & " AND Periodo = " & pPeriodo
    xCondicao = xCondicao & " AND [Numero da Ilha] = " & pIlha
    xCondicao = xCondicao & " AND Complemento = " & preparaTexto(pComplemento)
    xCondicao = xCondicao & " AND [Numero do Documento] = " & preparaTexto(pNumeroDocumento)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Movimento] ASC")
    If Localizar(1) Then
        LocalizarRegistroEspecialDoc = True
        AtribuiValor
    End If
    rsMovCaixaPista.Close
    Set rsMovCaixaPista = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarRegistroEspecialUsu(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pIlha As Integer, ByVal pComplemento As String, ByVal pNumeroConta As String, ByVal pDebitoCredito As String, ByVal pCodigoUsuario As Integer) As Boolean
    Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarRegistroEspecialUsu = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    If pDebitoCredito = "D" Then
        xCondicao = xCondicao & " AND [Numero da Conta Debito] = " & preparaTexto(pNumeroConta)
    Else
        xCondicao = xCondicao & " AND [Numero da Conta Credito] = " & preparaTexto(pNumeroConta)
    End If
    xCondicao = xCondicao & " AND Periodo = " & pPeriodo
    xCondicao = xCondicao & " AND [Numero da Ilha] = " & pIlha
    xCondicao = xCondicao & " AND Complemento = " & preparaTexto(pComplemento)
    xCondicao = xCondicao & " AND [Codigo do Usuario] = " & pCodigoUsuario
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Movimento] ASC")
    If Localizar(1) Then
        LocalizarRegistroEspecialUsu = True
        AtribuiValor
    End If
    rsMovCaixaPista.Close
    Set rsMovCaixaPista = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
'localiza o movimento do caixa referente ao juros sobre cheque pre-datado de um determinado movimento cheque
Public Function LocalizarMovimentoJurosMovimentoChequePreDatado(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pNumeroMovimentoCheque As Long, ByVal pNumeroDoDocumento As String) As Boolean
    Dim xCondicao As String

On Error GoTo trata_erro

        LocalizarMovimentoJurosMovimentoChequePreDatado = False
        
        xCondicao = " WHERE Empresa = " & pEmpresa
        xCondicao = xCondicao & " AND Data = " & preparaData(pData)
        xCondicao = xCondicao & " AND Periodo = " & pPeriodo
        xCondicao = xCondicao & " AND [Dados Interno] = " & preparaTexto("REF_MOV_CHEQUE|@|" & pNumeroMovimentoCheque)
        xCondicao = xCondicao & " AND [Numero do Documento] = " & preparaTexto(pNumeroDoDocumento)
        gSQL = PreparaSQL(xCondicao, " ORDER BY [Numero do Movimento] ASC")
        
        If Localizar(1) Then
            LocalizarMovimentoJurosMovimentoChequePreDatado = True
            AtribuiValor
        End If
    
    rsMovCaixaPista.Close
    Set rsMovCaixaPista = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data < " & preparaData(mData)
    xCondicao = xCondicao & " AND [Numero do Movimento] < " & mNumeroMovimento
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, [Numero do Movimento] DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsMovCaixaPista.Close
    Set rsMovCaixaPista = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal pNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL(" WHERE Nome = " & preparaTexto(pNome), "")
    LocalizarNome = Localizar(1)
    If LocalizarNome = True Then
        AtribuiValor
    End If
    rsMovCaixaPista.Close
    Set rsMovCaixaPista = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, [Numero do Movimento] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsMovCaixaPista.Close
    Set rsMovCaixaPista = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data > " & preparaData(mData)
    xCondicao = xCondicao & " AND [Numero do Movimento] > " & mNumeroMovimento
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, [Numero do Movimento] ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsMovCaixaPista.Close
    Set rsMovCaixaPista = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumeroMovimento As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(mData)
    xCondicao = xCondicao & " AND [Numero do Movimento] = " & pNumeroMovimento
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Data DESC, [Numero do Movimento] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsMovCaixaPista.Close
    Set rsMovCaixaPista = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    mNumeroMovimento = ProximoRegistro(mEmpresa, mData)
    gSQL = "INSERT INTO MovimentoCaixaPista ( Empresa, Data, [Numero do Movimento], Valor, [Numero do Documento], "
    gSQL = gSQL & "[Codigo do Historico], Complemento, [Numero da Conta Debito], [Numero da Conta Credito], "
    gSQL = gSQL & "[Tipo do Movimento], [Codigo do Usuario], Periodo, [Numero da Ilha], [Dados Interno], "
    gSQL = gSQL & "[Codigo do Lancamento Padrao], [Data da Digitacao], [Hora da Digitacao], "
    gSQL = gSQL & "[Data da Alteracao], [Hora da Alteracao] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlNumero(1, mNumeroMovimento, ", ")
    Call sqlValor(1, mValor, ", ")
    Call sqlTexto(1, mNumeroDocumento, ", ")
    Call sqlNumero(1, mCodigoHistorico, ", ")
    Call sqlTexto(1, mComplemento, ", ")
    Call sqlTexto(1, mNumeroContaDebito, ", ")
    Call sqlTexto(1, mNumeroContaCredito, ", ")
    Call sqlNumero(1, mTipoMovimento, ", ")
    Call sqlNumero(1, mCodigoUsuario, ", ")
    Call sqlNumero(1, mPeriodo, ", ")
    Call sqlNumero(1, mNumeroIlha, ", ")
    Call sqlTexto(1, mDadosInterno, ", ")
    Call sqlNumero(1, mCodigoLancamentoPadrao, ", ")
    Call sqlData(1, mDataDigitacao, ", ")
    Call sqlHora(1, mHoraDigitacao, ", ")
    Call sqlData(1, mDataAlteracao, ", ")
    Call sqlHora(1, mHoraAlteracao, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumeroMovimento As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE MovimentoCaixaPista SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlNumero(2, ", [Numero do Movimento] = ", mNumeroMovimento)
    Call sqlValor(2, ", Valor = ", mValor)
    Call sqlTexto(2, ", [Numero do Documento] = ", mNumeroDocumento)
    Call sqlNumero(2, ", [Codigo do Historico] = ", mCodigoHistorico)
    Call sqlTexto(2, ", Complemento = ", mComplemento)
    Call sqlTexto(2, ", [Numero da Conta Debito] = ", mNumeroContaDebito)
    Call sqlTexto(2, ", [Numero da Conta Credito] = ", mNumeroContaCredito)
    Call sqlNumero(2, ", [Tipo do Movimento] = ", mTipoMovimento)
    Call sqlNumero(2, ", [Codigo do Usuario] = ", mCodigoUsuario)
    Call sqlNumero(2, ", Periodo = ", mPeriodo)
    Call sqlNumero(2, ", [Numero da Ilha] = ", mNumeroIlha)
    Call sqlTexto(2, ", [Dados Interno] = ", mDadosInterno)
    Call sqlNumero(2, ", [Codigo do Lancamento Padrao] = ", mCodigoLancamentoPadrao)
    Call sqlData(2, ", [Data da Digitacao] = ", mDataDigitacao)
    Call sqlHora(2, ", [Hora da Digitacao] = ", mHoraDigitacao)
    Call sqlData(2, ", [Data da Alteracao] = ", mDataAlteracao)
    Call sqlHora(2, ", [Hora da Alteracao] = ", mHoraAlteracao)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Numero do Movimento] = " & pNumeroMovimento
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function

'Altera o subcaixa 999 para um novo
'específicamente no Movimento de combustivel (7)
Public Function AlteraSubCaixaMovBomba(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pNovoSubCaixa As Integer) As Boolean
Dim lRecordsAffected As Long
Dim xComplementoNovo
Dim xDadosInternoNovo

On Error GoTo TrataErro

    AlteraSubCaixaMovBomba = False
    xComplementoNovo = Replace(mComplemento, "SubCx:999", "SubCx:" & Format(pNovoSubCaixa, "000"))
    xDadosInternoNovo = "BOMBA|@|" & Format(pNovoSubCaixa, "000") & "|@|"
    
    gSQL = "UPDATE MovimentoCaixaPista SET "
    Call sqlTexto(2, "Complemento = ", xComplementoNovo)
    Call sqlTexto(2, ", [Dados Interno] = ", xDadosInternoNovo)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Periodo = " & pPeriodo
    gSQL = gSQL & " AND [Codigo do Lancamento Padrao] = " & 7
    gSQL = gSQL & " AND [Dados Interno] = " & preparaTexto("BOMBA|@|999|@|")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlteraSubCaixaMovBomba = True
    End If
    Exit Function

TrataErro:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumeroMovimento As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM MovimentoCaixaPista"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Numero do Movimento] = " & pNumeroMovimento
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, [Numero do Movimento], Valor, [Numero do Documento], [Codigo do Historico], Complemento, [Numero da Conta Debito], [Numero da Conta Credito], [Tipo do Movimento], [Codigo do Usuario], Periodo, [Numero da Ilha], [Dados Interno], [Codigo do Lancamento Padrao], [Data da Digitacao], [Hora da Digitacao], [Data da Alteracao], [Hora da Digitacao] FROM MovimentoCaixaPista" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovCaixaPista = New adodb.Recordset
    rsMovCaixaPista.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovCaixaPista.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCaixaPista.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovCaixaPista("Empresa").Value
    Data = rsMovCaixaPista("Data").Value
    NumeroMovimento = rsMovCaixaPista("Numero do Movimento").Value
    Valor = rsMovCaixaPista("Valor").Value
    NumeroDocumento = rsMovCaixaPista("Numero do Documento").Value
    CodigoHistorico = rsMovCaixaPista("Codigo do Historico").Value
    Complemento = rsMovCaixaPista("Complemento").Value
    NumeroContaDebito = rsMovCaixaPista("Numero da Conta Debito").Value
    NumeroContaCredito = rsMovCaixaPista("Numero da Conta Credito").Value
    TipoMovimento = rsMovCaixaPista("Tipo do Movimento").Value
    CodigoUsuario = rsMovCaixaPista("Codigo do Usuario").Value
    Periodo = rsMovCaixaPista("Periodo").Value
    NumeroIlha = rsMovCaixaPista("Numero da Ilha").Value
    DadosInterno = rsMovCaixaPista("Dados Interno").Value
    CodigoLancamentoPadrao = rsMovCaixaPista("Codigo do Lancamento Padrao").Value
    DataDigitacao = rsMovCaixaPista("Data da Digitacao").Value
    HoraDigitacao = rsMovCaixaPista("Hora da Digitacao").Value
    DataAlteracao = rsMovCaixaPista("Data da Digitacao").Value
    HoraAlteracao = rsMovCaixaPista("Hora da Digitacao").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function


Public Function ProximoRegistro(ByVal pEmpresa As Integer, ByVal pData As Date) As Long

On Error GoTo Err_ProximoRegistro

    'Pega número do último movimento
    ProximoRegistro = 1
    gSQL = ""
    gSQL = gSQL & "SELECT [Numero do Movimento]"
    gSQL = gSQL & "  FROM MovimentoCaixaPista"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Data = " & preparaData(pData)
    gSQL = gSQL & " ORDER BY [Numero do Movimento] DESC"
    If Localizar(1) Then
        ProximoRegistro = rsMovCaixaPista("Numero do Movimento").Value + 1
    End If
    rsMovCaixaPista.Close
    Set rsMovCaixaPista = Nothing
    Exit Function

Err_ProximoRegistro:
End Function

