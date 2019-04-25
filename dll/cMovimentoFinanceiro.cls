VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoFinanceiro"
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
Private mCodigoPortador As Integer
Private mCodigoUsuario As Integer
Private mPeriodo As Integer
Private mDadosInterno As String
Private mCodigoLancamentoFinanceiro As Integer
Private mDataDigitacao As Date
Private mHoraDigitacao As Date
Private mDataAlteracao As Date
Private mHoraAlteracao As Date

Private rsMovFinanceiro As New adodb.Recordset


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


Public Property Get CodigoPortador() As Integer
    CodigoPortador = mCodigoPortador
End Property
Public Property Let CodigoPortador(ByVal Valor As Integer)
    mCodigoPortador = Valor
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


Public Property Get DadosInterno() As String
    DadosInterno = mDadosInterno
End Property
Public Property Let DadosInterno(ByVal Valor As String)
    mDadosInterno = Valor
End Property


Public Property Get CodigoLancamentoFinanceiro() As Integer
    CodigoLancamentoFinanceiro = mCodigoLancamentoFinanceiro
End Property
Public Property Let CodigoLancamentoFinanceiro(ByVal Valor As Integer)
    mCodigoLancamentoFinanceiro = Valor
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
    rsMovFinanceiro.Close
    Set rsMovFinanceiro = Nothing
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
    rsMovFinanceiro.Close
    Set rsMovFinanceiro = Nothing
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
    xCondicao = xCondicao & " AND Complemento = " & preparaTexto(pComplemento)
    xCondicao = xCondicao & " AND [Numero do Documento] = " & preparaTexto(pNumeroDocumento)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Movimento] ASC")
    If Localizar(1) Then
        LocalizarRegistroEspecialDoc = True
        AtribuiValor
    End If
    rsMovFinanceiro.Close
    Set rsMovFinanceiro = Nothing
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
    rsMovFinanceiro.Close
    Set rsMovFinanceiro = Nothing
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
    rsMovFinanceiro.Close
    Set rsMovFinanceiro = Nothing
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
    rsMovFinanceiro.Close
    Set rsMovFinanceiro = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, [Numero do Movimento] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsMovFinanceiro.Close
    Set rsMovFinanceiro = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    mNumeroMovimento = ProximoRegistro(mEmpresa, mData)
    gSQL = "INSERT INTO MovimentoFinanceiro ( Empresa, Data, [Numero do Movimento], Valor, [Numero do Documento], "
    gSQL = gSQL & "[Codigo do Historico], Complemento, [Numero da Conta Debito], [Numero da Conta Credito], [Codigo do Portador], "
    gSQL = gSQL & "[Codigo do Usuario], Periodo, [Dados Interno], [Codigo do Lancamento Financeiro], [Data da Digitacao], "
    gSQL = gSQL & "[Hora da Digitacao], [Data da Alteracao], [Hora da Alteracao] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlNumero(1, mNumeroMovimento, ", ")
    Call sqlValor(1, mValor, ", ")
    Call sqlTexto(1, mNumeroDocumento, ", ")
    Call sqlNumero(1, mCodigoHistorico, ", ")
    Call sqlTexto(1, mComplemento, ", ")
    Call sqlTexto(1, mNumeroContaDebito, ", ")
    Call sqlTexto(1, mNumeroContaCredito, ", ")
    Call sqlNumero(1, mCodigoPortador, ", ")
    Call sqlNumero(1, mCodigoUsuario, ", ")
    Call sqlNumero(1, mPeriodo, ", ")
    Call sqlTexto(1, mDadosInterno, ", ")
    Call sqlNumero(1, mCodigoLancamentoFinanceiro, ", ")
    Call sqlData(1, mDataDigitacao, ", ")
    Call sqlHora(1, mHoraDigitacao, ", ")
    Call sqlData(1, mDataAlteracao, ", ")
    Call sqlHora(1, mHoraAlteracao, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
        If Not LoopIncluiSaldo() Then
            Call CriaLogCadastroDll2("MovimentoFinanceiro:Incluir - Erro na atualização do saldo.", Err.Description, gSQL)
        End If
    Else
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumeroMovimento As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    If Not LoopExcluiSaldo(pEmpresa, pData, pNumeroMovimento) Then
        Exit Function
    End If
    gSQL = "UPDATE MovimentoFinanceiro SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlNumero(2, ", [Numero do Movimento] = ", mNumeroMovimento)
    Call sqlValor(2, ", Valor = ", mValor)
    Call sqlTexto(2, ", [Numero do Documento] = ", mNumeroDocumento)
    Call sqlNumero(2, ", [Codigo do Historico] = ", mCodigoHistorico)
    Call sqlTexto(2, ", Complemento = ", mComplemento)
    Call sqlTexto(2, ", [Numero da Conta Debito] = ", mNumeroContaDebito)
    Call sqlTexto(2, ", [Numero da Conta Credito] = ", mNumeroContaCredito)
    Call sqlNumero(2, ", [Codigo do Portador] = ", mCodigoPortador)
    Call sqlNumero(2, ", [Codigo do Usuario] = ", mCodigoUsuario)
    Call sqlNumero(2, ", Periodo = ", mPeriodo)
    Call sqlTexto(2, ", [Dados Interno] = ", mDadosInterno)
    Call sqlNumero(2, ", [Codigo do Lancamento Financeiro] = ", mCodigoLancamentoFinanceiro)
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
        If Not LoopIncluiSaldo() Then
            Call CriaLogCadastroDll2("MovimentoFinanceiro:Alterar - Erro na atualização do saldo.", Err.Description, gSQL)
        End If
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumeroMovimento As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    If Not LoopExcluiSaldo(pEmpresa, pData, pNumeroMovimento) Then
        Exit Function
    End If
    gSQL = "DELETE FROM MovimentoFinanceiro"
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
    PreparaSQL = Trim("SELECT Empresa, Data, [Numero do Movimento], Valor, [Numero do Documento], [Codigo do Historico], Complemento, [Numero da Conta Debito], [Numero da Conta Credito], [Codigo do Portador], [Codigo do Usuario], Periodo, [Dados Interno], [Codigo do Lancamento Financeiro], [Data da Digitacao], [Hora da Digitacao], [Data da Alteracao], [Hora da Alteracao] FROM MovimentoFinanceiro" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovFinanceiro = New adodb.Recordset
    rsMovFinanceiro.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovFinanceiro.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovFinanceiro.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovFinanceiro("Empresa").Value
    Data = rsMovFinanceiro("Data").Value
    NumeroMovimento = rsMovFinanceiro("Numero do Movimento").Value
    Valor = rsMovFinanceiro("Valor").Value
    NumeroDocumento = rsMovFinanceiro("Numero do Documento").Value
    CodigoHistorico = rsMovFinanceiro("Codigo do Historico").Value
    Complemento = rsMovFinanceiro("Complemento").Value
    NumeroContaDebito = rsMovFinanceiro("Numero da Conta Debito").Value
    NumeroContaCredito = rsMovFinanceiro("Numero da Conta Credito").Value
    CodigoPortador = rsMovFinanceiro("Codigo do Portador").Value
    CodigoUsuario = rsMovFinanceiro("Codigo do Usuario").Value
    Periodo = rsMovFinanceiro("Periodo").Value
    DadosInterno = rsMovFinanceiro("Dados Interno").Value
    CodigoLancamentoFinanceiro = rsMovFinanceiro("Codigo do Lancamento Financeiro").Value
    DataDigitacao = rsMovFinanceiro("Data da Digitacao").Value
    HoraDigitacao = rsMovFinanceiro("Hora da Digitacao").Value
    If IsNull(rsMovFinanceiro("Data da Alteracao").Value) Then
        DataAlteracao = "00:00:00"
    Else
        DataAlteracao = rsMovFinanceiro("Data da Alteracao").Value
    End If
    If IsNull(rsMovFinanceiro("Hora da Alteracao").Value) Then
        HoraAlteracao = "00:00:00"
    Else
        HoraAlteracao = rsMovFinanceiro("Hora da Alteracao").Value
    End If
End Sub

Private Function ExcluiSaldo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumeroConta As String, ByVal pDebitoCredito As String, ByVal pValor As Currency) As Boolean
    Dim lRecordsAffected As Long
    
    ExcluiSaldo = False

    If Mid(pNumeroConta, 1, 1) = "2" Or Mid(pNumeroConta, 1, 1) = "4" Then
        If pDebitoCredito = "C" Then
            pValor = -pValor
        Else
            pValor = pValor
        End If
    Else
        If pDebitoCredito = "D" Then
            pValor = pValor
        Else
            pValor = -pValor
        End If
    End If

    'Atualiza Saldo com Data >= a Atual
    gSQL = "UPDATE Saldo_Conta SET "
    Call sqlValor(2, "Saldo = Saldo - ", CStr(pValor))
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND [Codigo da Conta] = " & preparaTexto(pNumeroConta)
    gSQL = gSQL & "   AND Data >= " & preparaData(pData)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        ExcluiSaldo = True
    Else
        Call CriaLogCadastroDll2("MovimentoFinanceiro:ExcluiSaldo - Erro ao alterar saldo.", Err.Description, gSQL)
    End If
End Function

Private Function IncluiSaldo(ByVal pCodigoConta As String, ByVal pDebitoCredito As String) As Boolean
    Dim lRecordsAffected As Long
    Dim xSaldo As Currency
    Dim xValor As Currency
    Dim xOperador As String

On Error GoTo Err_IncluiSaldo
    
    IncluiSaldo = False
    If Mid(pCodigoConta, 1, 1) = "2" Or Mid(pCodigoConta, 1, 1) = "4" Then
        If pDebitoCredito = "C" Then
            xValor = -mValor
        Else
            xValor = mValor
        End If
    Else
        If pDebitoCredito = "D" Then
            xValor = mValor
        Else
            xValor = -mValor
        End If
    End If

    'Busca Saldo Anterior a Data Atual
    xSaldo = 0
    gSQL = ""
    gSQL = gSQL & "SELECT Saldo"
    gSQL = gSQL & "  FROM Saldo_Conta"
    gSQL = gSQL & " WHERE Empresa = " & mEmpresa
    gSQL = gSQL & "   AND [Codigo da Conta] = " & preparaTexto(pCodigoConta)
    gSQL = gSQL & "   AND Data < " & preparaData(mData)
    gSQL = gSQL & " ORDER BY Data DESC"
    If Localizar(1) Then
        xSaldo = rsMovFinanceiro("Saldo").Value
    End If
    rsMovFinanceiro.Close
    Set rsMovFinanceiro = Nothing

    'Busca Saldo na Data Atual
    gSQL = ""
    gSQL = gSQL & "SELECT Saldo"
    gSQL = gSQL & "  FROM Saldo_Conta"
    gSQL = gSQL & " WHERE Empresa = " & mEmpresa
    gSQL = gSQL & "   AND [Codigo da Conta] = " & preparaTexto(pCodigoConta)
    gSQL = gSQL & "   AND Data = " & preparaData(mData)
    gSQL = gSQL & " ORDER BY Data DESC"
    If Localizar(1) Then
        xOperador = ">="
        xSaldo = rsMovFinanceiro("Saldo").Value
        rsMovFinanceiro.Close
        Set rsMovFinanceiro = Nothing
    Else
        xOperador = ">="
        rsMovFinanceiro.Close
        Set rsMovFinanceiro = Nothing
        'Inclui Saldo na Data Atual
        'xSaldo = xSaldo + xValor
        gSQL = "INSERT INTO Saldo_Conta ( Empresa, [Codigo da Conta], Data, Saldo ) VALUES ( "
        Call sqlNumero(1, CStr(mEmpresa), ", ")
        Call sqlTexto(1, CStr(pCodigoConta), ", ")
        Call sqlData(1, CStr(mData), ", ")
        Call sqlValor(1, CStr(xSaldo), " )")
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        If lRecordsAffected = 0 Then
            Call CriaLogCadastroDll2("MovimentoFinanceiro:IncluiSaldo - Erro ao incluir saldo.", Err.Description, gSQL)
            Exit Function
        End If
    End If

    'Atualiza Saldo com Data >= ou = a Atual
    'Conforme variavel xOperador
    gSQL = "UPDATE Saldo_Conta SET "
    Call sqlValor(2, "Saldo = Saldo + ", CStr(xValor))
    gSQL = gSQL & " WHERE Empresa = " & mEmpresa
    gSQL = gSQL & "   AND [Codigo da Conta] = " & preparaTexto(pCodigoConta)
    gSQL = gSQL & "   AND Data " & xOperador & " " & preparaData(mData)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        IncluiSaldo = True
    Else
        Call CriaLogCadastroDll2("MovimentoFinanceiro:IncluiSaldo - Erro ao alterar saldo.", Err.Description, gSQL)
    End If
    Exit Function

Err_IncluiSaldo:
    Call CriaLogCadastroDll2("MovimentoFinanceiro:IncluiSaldo - Erro desconhecido.", Err.Description, gSQL)
End Function

Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

Private Function LoopExcluiSaldo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumeroMovimento As Integer) As Boolean
    Dim xContaDebito As String
    Dim xContaCredito As String
    Dim xValor As Currency

    LoopExcluiSaldo = False
    xContaDebito = ""
    xContaCredito = ""
    xValor = 0

    'Busca registro atual do Movimento_Caixa
    gSQL = ""
    gSQL = gSQL & "SELECT Data, Valor, [Numero da Conta Debito], [Numero da Conta Credito]"
    gSQL = gSQL & "  FROM MovimentoFinanceiro"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Data = " & preparaData(pData)
    gSQL = gSQL & "   AND [Numero do Movimento] = " & pNumeroMovimento
    If Localizar(1) Then
        xContaDebito = rsMovFinanceiro("Numero da Conta Debito").Value
        xContaCredito = rsMovFinanceiro("Numero da Conta Credito").Value
        xValor = rsMovFinanceiro("Valor").Value
    End If
    rsMovFinanceiro.Close
    Set rsMovFinanceiro = Nothing

    If xContaDebito <> "" Then
        LoopExcluiSaldo = ExcluiSaldo(pEmpresa, pData, xContaDebito, "C", xValor)
    End If
    If xContaCredito <> "" Then
        LoopExcluiSaldo = ExcluiSaldo(pEmpresa, pData, xContaCredito, "D", xValor)
    End If
End Function

Private Function LoopIncluiSaldo() As Boolean
    LoopIncluiSaldo = False
    If mNumeroContaDebito <> "" Then
        LoopIncluiSaldo = IncluiSaldo(mNumeroContaDebito, "C")
    End If
    If mNumeroContaCredito <> "" Then
        LoopIncluiSaldo = IncluiSaldo(mNumeroContaCredito, "D")
    End If
End Function

Public Function ProximoRegistro(ByVal pEmpresa As Integer, ByVal pData As Date) As Long

On Error GoTo Err_ProximoRegistro

    'Pega número do último movimento
    ProximoRegistro = 1
    gSQL = ""
    gSQL = gSQL & "SELECT [Numero do Movimento]"
    gSQL = gSQL & "  FROM MovimentoFinanceiro"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Data = " & preparaData(pData)
    gSQL = gSQL & " ORDER BY [Numero do Movimento] DESC"
    If Localizar(1) Then
        ProximoRegistro = rsMovFinanceiro("Numero do Movimento").Value + 1
    End If
    rsMovFinanceiro.Close
    Set rsMovFinanceiro = Nothing
    Exit Function

Err_ProximoRegistro:
End Function

