VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSuprimentoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mData As Date
Private mPeriodo As Integer
Private mNumeroIlha As Integer
Private mCodigoComposicao As Integer
Private mValor As Currency

Private rsSuprimentoCaixa As New adodb.Recordset


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


Public Property Get CodigoComposicao() As Integer
    CodigoComposicao = mCodigoComposicao
End Property
Public Property Let CodigoComposicao(ByVal Valor As Integer)
    mCodigoComposicao = Valor
End Property


Public Property Get Valor() As Currency
    Valor = mValor
End Property
Public Property Let Valor(ByVal Valor As Currency)
    mValor = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pNumeroIlha As Integer, ByVal pCodigoComposicao As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND Periodo = " & pPeriodo
    xCondicao = xCondicao & " AND [Numero da Ilha] = " & pNumeroIlha
    xCondicao = xCondicao & " AND [Codigo da Composicao] = " & pCodigoComposicao
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsSuprimentoCaixa.Close
    Set rsSuprimentoCaixa = Nothing
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
    xCondicao = xCondicao & " AND Periodo < " & mPeriodo
    xCondicao = xCondicao & " AND [Numero da Ilha] < " & mNumeroIlha
    xCondicao = xCondicao & " AND [Codigo da Composicao] < " & mCodigoComposicao
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Data DESC, Periodo DESC, [Numero da Ilha] DESC, [Codigo da Composicao] DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsSuprimentoCaixa.Close
    Set rsSuprimentoCaixa = Nothing
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
    rsSuprimentoCaixa.Close
    Set rsSuprimentoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data > " & preparaData(mData)
    xCondicao = xCondicao & " AND Periodo > " & mPeriodo
    xCondicao = xCondicao & " AND [Numero da Ilha] > " & mNumeroIlha
    xCondicao = xCondicao & " AND [Codigo da Composicao] > " & mCodigoComposicao
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Data ASC, Periodo ASC, [Numero da Ilha] ASC, [Codigo da Composicao] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsSuprimentoCaixa.Close
    Set rsSuprimentoCaixa = Nothing
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
    xCondicao = xCondicao & " AND Periodo > " & mPeriodo
    xCondicao = xCondicao & " AND [Numero da Ilha] > " & mNumeroIlha
    xCondicao = xCondicao & " AND [Codigo da Composicao] > " & mCodigoComposicao
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Data ASC, Periodo ASC, [Numero da Ilha] ASC, [Codigo da Composicao] ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsSuprimentoCaixa.Close
    Set rsSuprimentoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pNumeroIlha As Integer, ByVal pCodigoComposicao As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(mData)
    xCondicao = xCondicao & " AND Periodo = " & pPeriodo
    xCondicao = xCondicao & " AND [Numero da Ilha] = " & pNumeroIlha
    xCondicao = xCondicao & " AND [Codigo da Composicao] = " & pCodigoComposicao
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Data DESC, Periodo DESC, [Numero da Ilha] DESC, [Codigo da Composicao] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsSuprimentoCaixa.Close
    Set rsSuprimentoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO SuprimentoCaixa ( Empresa, Data, Periodo, [Numero da Ilha], [Codigo da Composicao], "
    gSQL = gSQL & "Valor ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlNumero(1, mPeriodo, ", ")
    Call sqlNumero(1, mNumeroIlha, ", ")
    Call sqlNumero(1, mCodigoComposicao, ", ")
    Call sqlValor(1, mValor, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pNumeroIlha As Integer, ByVal pCodigoComposicao As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE SuprimentoCaixa SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlNumero(2, ", Periodo = ", mPeriodo)
    Call sqlNumero(2, ", [Numero da Ilha] = ", mNumeroIlha)
    Call sqlNumero(2, ", [Codigo da Composicao] = ", mCodigoComposicao)
    Call sqlValor(2, ", Valor = ", mValor)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Periodo = " & pPeriodo
    gSQL = gSQL & " AND [Numero da Ilha] = " & pNumeroIlha
    gSQL = gSQL & " AND [Codigo da Composicao] = " & pCodigoComposicao
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pNumeroIlha As Integer, ByVal pCodigoComposicao As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM SuprimentoCaixa"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Periodo = " & pPeriodo
    gSQL = gSQL & " AND [Numero da Ilha] = " & pNumeroIlha
    gSQL = gSQL & " AND [Codigo da Composicao] = " & pCodigoComposicao
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


Public Function TotalPeriodo(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer) As Currency

On Error GoTo trata_erro
    TotalPeriodo = 0
    gSQL = "SELECT SUM(Valor) AS Total"
    gSQL = gSQL & " FROM SuprimentoCaixa"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND Data <= " & preparaData(pDataFinal)
    gSQL = gSQL & " AND Periodo >= " & pPeriodoInicial
    gSQL = gSQL & " AND Periodo <= " & pPeriodoFinal
    If Localizar(0) Then
        If Not IsNull(rsSuprimentoCaixa("Total").Value) Then
            TotalPeriodo = rsSuprimentoCaixa("Total").Value
        End If
    End If
    rsSuprimentoCaixa.Close
    Set rsSuprimentoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, Periodo, [Numero da Ilha], [Codigo da Composicao], Valor FROM SuprimentoCaixa" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsSuprimentoCaixa = New adodb.Recordset
    rsSuprimentoCaixa.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsSuprimentoCaixa.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsSuprimentoCaixa.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsSuprimentoCaixa("Empresa").Value
    Data = rsSuprimentoCaixa("Data").Value
    Periodo = rsSuprimentoCaixa("Periodo").Value
    NumeroIlha = rsSuprimentoCaixa("Numero da Ilha").Value
    CodigoComposicao = rsSuprimentoCaixa("Codigo da Composicao").Value
    Valor = rsSuprimentoCaixa("Valor").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

