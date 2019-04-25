VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAberturaCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mDataAbertura As Date
Private mTipoCaixa As String
Private mPeriodo As Integer
Private mNumeroIlha As Integer
Private mCodigoFuncionario As Integer
Private mHoraAbertura As Date
Private mDataFechamento As Date
Private mHoraFechamento As Date
Private mTipoMovimento As Integer
Private mFechadoPeloNivel As Integer
Private mRecebidoPeloFinanceiro As Boolean
Private mDataConferencia As Date
Private mConferidopeloNivel As Integer

Private rsAberturaCaixa As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get DataAbertura() As Date
    DataAbertura = mDataAbertura
End Property
Public Property Let DataAbertura(ByVal Valor As Date)
    mDataAbertura = Valor
End Property


Public Property Get TipoCaixa() As String
    TipoCaixa = mTipoCaixa
End Property
Public Property Let TipoCaixa(ByVal Valor As String)
    mTipoCaixa = Valor
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


Public Property Get CodigoFuncionario() As Integer
    CodigoFuncionario = mCodigoFuncionario
End Property
Public Property Let CodigoFuncionario(ByVal Valor As Integer)
    mCodigoFuncionario = Valor
End Property


Public Property Get HoraAbertura() As Date
    HoraAbertura = mHoraAbertura
End Property
Public Property Let HoraAbertura(ByVal Valor As Date)
    mHoraAbertura = Valor
End Property


Public Property Get DataFechamento() As Date
    DataFechamento = mDataFechamento
End Property
Public Property Let DataFechamento(ByVal Valor As Date)
    mDataFechamento = Valor
End Property


Public Property Get HoraFechamento() As Date
    HoraFechamento = mHoraFechamento
End Property
Public Property Let HoraFechamento(ByVal Valor As Date)
    mHoraFechamento = Valor
End Property


Public Property Get TipoMovimento() As Integer
    TipoMovimento = mTipoMovimento
End Property
Public Property Let TipoMovimento(ByVal Valor As Integer)
    mTipoMovimento = Valor
End Property


Public Property Get FechadoPeloNivel() As Integer
    FechadoPeloNivel = mFechadoPeloNivel
End Property
Public Property Let FechadoPeloNivel(ByVal Valor As Integer)
    mFechadoPeloNivel = Valor
End Property


Public Property Get RecebidoPeloFinanceiro() As Boolean
    RecebidoPeloFinanceiro = mRecebidoPeloFinanceiro
End Property
Public Property Let RecebidoPeloFinanceiro(ByVal Valor As Boolean)
    mRecebidoPeloFinanceiro = Valor
End Property


Public Property Get DataConferencia() As Date
    DataConferencia = mDataConferencia
End Property
Public Property Let DataConferencia(ByVal Valor As Date)
    mDataConferencia = Valor
End Property


Public Property Get ConferidopeloNivel() As Integer
    ConferidopeloNivel = mConferidopeloNivel
End Property
Public Property Let ConferidopeloNivel(ByVal Valor As Integer)
    mConferidopeloNivel = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe


Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pDataAbertura As Date, ByVal pTipoCaixa As String, ByVal pPeriodo As Integer, ByVal pNumeroIlha As Integer, ByVal pCodigoFuncionario As Integer, ByVal pTipoMovimento As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Data da Abertura] = " & preparaData(pDataAbertura)
    xCondicao = xCondicao & " AND [Tipo de Caixa] = " & preparaTexto(pTipoCaixa)
    xCondicao = xCondicao & " AND Periodo = " & pPeriodo
    xCondicao = xCondicao & " AND [Numero da Ilha] = " & pNumeroIlha
    xCondicao = xCondicao & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    xCondicao = xCondicao & " AND [Tipo do Movimento] = " & pTipoMovimento
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsAberturaCaixa.Close
    Set rsAberturaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarCxData(ByVal pEmpresa As Integer, ByVal pDataAbertura As Date, ByVal pTipoCaixa As String, ByVal pPeriodo As Integer, ByVal pNumeroIlha As Integer, ByVal pTipoMovimento As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCxData = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Data da Abertura] = " & preparaData(pDataAbertura)
    xCondicao = xCondicao & " AND [Tipo de Caixa] = " & preparaTexto(pTipoCaixa)
    xCondicao = xCondicao & " AND Periodo = " & pPeriodo
    xCondicao = xCondicao & " AND [Numero da Ilha] = " & pNumeroIlha
    xCondicao = xCondicao & " AND [Tipo do Movimento] = " & pTipoMovimento
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        LocalizarCxData = True
        AtribuiValor
    End If
    rsAberturaCaixa.Close
    Set rsAberturaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarCxDataFunc(ByVal pEmpresa As Integer, ByVal pDataAbertura As Date, ByVal pTipoCaixa As String, ByVal pPeriodo As Integer, ByVal pNumeroIlha As Integer, ByVal pTipoMovimento As Integer, ByVal pCodigoFuncionario As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCxDataFunc = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Data da Abertura] = " & preparaData(pDataAbertura)
    xCondicao = xCondicao & " AND [Tipo de Caixa] = " & preparaTexto(pTipoCaixa)
    xCondicao = xCondicao & " AND [Numero da Ilha] = " & pNumeroIlha
    xCondicao = xCondicao & " AND [Tipo do Movimento] = " & pTipoMovimento
    xCondicao = xCondicao & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        LocalizarCxDataFunc = True
        AtribuiValor
    End If
    rsAberturaCaixa.Close
    Set rsAberturaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Data da Abertura] < " & preparaData(mDataAbertura)
    xCondicao = xCondicao & " AND [Tipo de Caixa] < " & preparaTexto(mTipoCaixa)
    xCondicao = xCondicao & " AND Periodo < " & mPeriodo
    xCondicao = xCondicao & " AND [Numero da Ilha] < " & mNumeroIlha
    xCondicao = xCondicao & " AND [Codigo do Funcionario] < " & mCodigoFuncionario
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, [Data da Abertura] DESC, [Tipo de Caixa] DESC, Periodo DESC, [Numero da Ilha] DESC, [Codigo do Funcionario] DESC, [Tipo do Movimento] DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsAberturaCaixa.Close
    Set rsAberturaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data da Abertura] ASC, [Tipo de Caixa] ASC, Periodo ASC, [Numero da Ilha] ASC, [Codigo do Funcionario] ASC, [Tipo do Movimento] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsAberturaCaixa.Close
    Set rsAberturaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Data da Abertura] > " & preparaData(mDataAbertura)
    xCondicao = xCondicao & " AND [Tipo de Caixa] > " & preparaTexto(mTipoCaixa)
    xCondicao = xCondicao & " AND Periodo > " & mPeriodo
    xCondicao = xCondicao & " AND [Numero da Ilha] > " & mNumeroIlha
    xCondicao = xCondicao & " AND [Codigo do Funcionario] > " & mCodigoFuncionario
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, [Data da Abertura] ASC, [Tipo de Caixa] ASC, Periodo ASC, [Numero da Ilha] ASC, [Codigo do Funcionario] ASC, [Tipo do Movimento] ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsAberturaCaixa.Close
    Set rsAberturaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltAbertoDataFunc(ByVal pEmpresa As Integer, ByVal pDataAbertura As Date, ByVal pTipoCaixa As String, ByVal pNumeroIlha As Integer, ByVal pTipoMovimento As Integer, ByVal pCodigoFuncionario As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltAbertoDataFunc = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Data da Abertura] = " & preparaData(pDataAbertura)
    xCondicao = xCondicao & " AND [Tipo de Caixa] = " & preparaTexto(pTipoCaixa)
    xCondicao = xCondicao & " AND [Numero da Ilha] = " & pNumeroIlha
    xCondicao = xCondicao & " AND [Tipo do Movimento] = " & pTipoMovimento
    xCondicao = xCondicao & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    xCondicao = xCondicao & " AND [Data do Fechamento] IS NULL"
    gSQL = PreparaSQL(xCondicao, "ORDER BY PERIODO DESC")
    If Localizar(1) Then
        LocalizarUltAbertoDataFunc = True
        AtribuiValor
    End If
    rsAberturaCaixa.Close
    Set rsAberturaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltFechadoDataFunc(ByVal pEmpresa As Integer, ByVal pDataAbertura As Date, ByVal pTipoCaixa As String, ByVal pNumeroIlha As Integer, ByVal pTipoMovimento As Integer, ByVal pCodigoFuncionario As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltFechadoDataFunc = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Data da Abertura] = " & preparaData(pDataAbertura)
    xCondicao = xCondicao & " AND [Tipo de Caixa] = " & preparaTexto(pTipoCaixa)
    xCondicao = xCondicao & " AND [Numero da Ilha] = " & pNumeroIlha
    xCondicao = xCondicao & " AND [Tipo do Movimento] = " & pTipoMovimento
    xCondicao = xCondicao & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    xCondicao = xCondicao & " AND NOT [Data do Fechamento] IS NULL"
    gSQL = PreparaSQL(xCondicao, "ORDER BY PERIODO DESC")
    If Localizar(1) Then
        LocalizarUltFechadoDataFunc = True
        AtribuiValor
    End If
    rsAberturaCaixa.Close
    Set rsAberturaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer, ByVal pDataAbertura As Date, ByVal pTipoCaixa As String, ByVal pPeriodo As Integer, ByVal pNumeroIlha As Integer, ByVal pCodigoFuncionario As Integer, ByVal pTipoMovimento As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data da Abertura] DESC, [Tipo de Caixa] DESC, Periodo DESC, [Numero da Ilha] DESC, [Codigo do Funcionario] DESC, [Tipo do Movimento] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsAberturaCaixa.Close
    Set rsAberturaCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO AberturaCaixa ( Empresa, [Data da Abertura], [Tipo de Caixa], Periodo, [Numero da Ilha], "
    gSQL = gSQL & "[Codigo do Funcionario], [Hora da Abertura], [Data do Fechamento], [Hora do Fechamento], [Tipo do Movimento], "
    gSQL = gSQL & "[Fechado pelo Nivel], [Recebido Pelo Financeiro], [Data da Conferencia], [Conferido pelo Nivel] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mDataAbertura, ", ")
    Call sqlTexto(1, mTipoCaixa, ", ")
    Call sqlNumero(1, mPeriodo, ", ")
    Call sqlNumero(1, mNumeroIlha, ", ")
    Call sqlNumero(1, mCodigoFuncionario, ", ")
    Call sqlHora(1, mHoraAbertura, ", ")
    Call sqlData(1, mDataFechamento, ", ")
    Call sqlHora(1, mHoraFechamento, ", ")
    Call sqlNumero(1, mTipoMovimento, ", ")
    Call sqlNumero(1, mFechadoPeloNivel, ", ")
    Call sqlBoolean(1, mRecebidoPeloFinanceiro, ", ")
    Call sqlData(1, mDataConferencia, ", ")
    Call sqlNumero(1, mConferidopeloNivel, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pDataAbertura As Date, ByVal pTipoCaixa As String, ByVal pPeriodo As Integer, ByVal pNumeroIlha As Integer, ByVal pCodigoFuncionario As Integer, ByVal pTipoMovimento As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE AberturaCaixa SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", [Data da Abertura] = ", mDataAbertura)
    Call sqlTexto(2, ", [Tipo de Caixa] = ", mTipoCaixa)
    Call sqlNumero(2, ", Periodo = ", mPeriodo)
    Call sqlNumero(2, ", [Numero da Ilha] = ", mNumeroIlha)
    Call sqlNumero(2, ", [Codigo do Funcionario] = ", mCodigoFuncionario)
    Call sqlHora(2, ", [Hora da Abertura] = ", mHoraAbertura)
    Call sqlData(2, ", [Data do Fechamento] = ", mDataFechamento)
    Call sqlHora(2, ", [Hora do Fechamento] = ", mHoraFechamento)
    Call sqlNumero(2, ", [Tipo do Movimento] = ", mTipoMovimento)
    Call sqlNumero(2, ", [Fechado pelo Nivel] = ", mFechadoPeloNivel)
    Call sqlBoolean(2, ", [Recebido Pelo Financeiro] = ", mRecebidoPeloFinanceiro)
    Call sqlData(2, ", [Data da Conferencia] = ", mDataConferencia)
    Call sqlNumero(2, ", [Conferido pelo Nivel] = ", mConferidopeloNivel)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Data da Abertura] = " & preparaData(pDataAbertura)
    gSQL = gSQL & " AND [Tipo de Caixa] = " & pTipoCaixa
    gSQL = gSQL & " AND Periodo = " & pPeriodo
    gSQL = gSQL & " AND [Numero da Ilha] = " & pNumeroIlha
    gSQL = gSQL & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    gSQL = gSQL & " AND [Tipo do Movimento] = " & pTipoMovimento
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pDataAbertura As Date, ByVal pTipoCaixa As String, ByVal pPeriodo As Integer, ByVal pNumeroIlha As Integer, ByVal pCodigoFuncionario As Integer, ByVal pTipoMovimento As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM AberturaCaixa"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Data da Abertura] = " & preparaData(pDataAbertura)
    gSQL = gSQL & " AND [Tipo de Caixa] = " & pTipoCaixa
    gSQL = gSQL & " AND Periodo = " & pPeriodo
    gSQL = gSQL & " AND [Numero da Ilha] = " & pNumeroIlha
    gSQL = gSQL & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    gSQL = gSQL & " AND [Tipo do Movimento] = " & pTipoMovimento
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Data da Abertura], [Tipo de Caixa], Periodo, [Numero da Ilha], [Codigo do Funcionario], [Hora da Abertura], [Data do Fechamento], [Hora do Fechamento], [Tipo do Movimento], [Fechado pelo Nivel], [Recebido Pelo Financeiro], [Data da Conferencia], [Conferido pelo Nivel] FROM AberturaCaixa" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsAberturaCaixa = New adodb.Recordset
    rsAberturaCaixa.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsAberturaCaixa.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsAberturaCaixa.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsAberturaCaixa("Empresa").Value
    DataAbertura = rsAberturaCaixa("Data da Abertura").Value
    TipoCaixa = rsAberturaCaixa("Tipo de Caixa").Value
    Periodo = rsAberturaCaixa("Periodo").Value
    NumeroIlha = rsAberturaCaixa("Numero da Ilha").Value
    CodigoFuncionario = rsAberturaCaixa("Codigo do Funcionario").Value
    HoraAbertura = rsAberturaCaixa("Hora da Abertura").Value
    If IsNull(rsAberturaCaixa("Data do Fechamento").Value) Then
        DataFechamento = "00:00:00"
    Else
        DataFechamento = rsAberturaCaixa("Data do Fechamento").Value
    End If
    If IsNull(rsAberturaCaixa("Hora do Fechamento").Value) Then
        HoraFechamento = "00:00:00"
    Else
        HoraFechamento = rsAberturaCaixa("Hora do Fechamento").Value
    End If
    TipoMovimento = rsAberturaCaixa("Tipo do Movimento").Value
    FechadoPeloNivel = rsAberturaCaixa("Fechado pelo Nivel").Value
    RecebidoPeloFinanceiro = rsAberturaCaixa("Recebido Pelo Financeiro").Value
    If IsNull(rsAberturaCaixa("Data da Conferencia").Value) Then
        DataConferencia = "00:00:00"
    Else
        DataConferencia = rsAberturaCaixa("Data da Conferencia").Value
    End If
    ConferidopeloNivel = rsAberturaCaixa("Conferido pelo Nivel").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

