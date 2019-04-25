VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDuplicataReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mNumeroDocumento As Long
Private mCodigoCliente As Long
Private mDataEmissao As Date
Private mDataPeriodoInicial As Date
Private mDataPeriodoFinal As Date
Private mNumeroDuplicata As String
Private mDataVencimento As Date
Private mValorCustoBancario As Currency
Private mValorVencimento As Currency
Private mNumeroMovimentoCaixa As Long
Private mNumeroContaBancaria As String

Private rsDuplicataReceber As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get NumeroDocumento() As Long
    NumeroDocumento = mNumeroDocumento
End Property
Public Property Let NumeroDocumento(ByVal Valor As Long)
    mNumeroDocumento = Valor
End Property


Public Property Get CodigoCliente() As Long
    CodigoCliente = mCodigoCliente
End Property
Public Property Let CodigoCliente(ByVal Valor As Long)
    mCodigoCliente = Valor
End Property


Public Property Get DataEmissao() As Date
    DataEmissao = mDataEmissao
End Property
Public Property Let DataEmissao(ByVal Valor As Date)
    mDataEmissao = Valor
End Property


Public Property Get DataPeriodoInicial() As Date
    DataPeriodoInicial = mDataPeriodoInicial
End Property
Public Property Let DataPeriodoInicial(ByVal Valor As Date)
    mDataPeriodoInicial = Valor
End Property


Public Property Get DataPeriodoFinal() As Date
    DataPeriodoFinal = mDataPeriodoFinal
End Property
Public Property Let DataPeriodoFinal(ByVal Valor As Date)
    mDataPeriodoFinal = Valor
End Property


Public Property Get NumeroDuplicata() As String
    NumeroDuplicata = mNumeroDuplicata
End Property
Public Property Let NumeroDuplicata(ByVal Valor As String)
    mNumeroDuplicata = Valor
End Property


Public Property Get DataVencimento() As Date
    DataVencimento = mDataVencimento
End Property
Public Property Let DataVencimento(ByVal Valor As Date)
    mDataVencimento = Valor
End Property


Public Property Get ValorCustoBancario() As Currency
    ValorCustoBancario = mValorCustoBancario
End Property
Public Property Let ValorCustoBancario(ByVal Valor As Currency)
    mValorCustoBancario = Valor
End Property


Public Property Get ValorVencimento() As Currency
    ValorVencimento = mValorVencimento
End Property
Public Property Let ValorVencimento(ByVal Valor As Currency)
    mValorVencimento = Valor
End Property


Public Property Get NumeroMovimentoCaixa() As Long
    NumeroMovimentoCaixa = mNumeroMovimentoCaixa
End Property
Public Property Let NumeroMovimentoCaixa(ByVal Valor As Long)
    mNumeroMovimentoCaixa = Valor
End Property


Public Property Get NumeroContaBancaria() As String
    NumeroContaBancaria = mNumeroContaBancaria
End Property
Public Property Let NumeroContaBancaria(ByVal Valor As String)
    mNumeroContaBancaria = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nNumeroDocumento As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Numero do Documento] = " & nNumeroDocumento
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsDuplicataReceber.Close
    Set rsDuplicataReceber = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNumeroDuplicata(ByVal nEmpresa As Integer, ByVal nNumeroDuplicata As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarNumeroDuplicata = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Numero da Duplicata] = " & preparaTexto(nNumeroDuplicata)
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        LocalizarNumeroDuplicata = True
        AtribuiValor
    End If
    rsDuplicataReceber.Close
    Set rsDuplicataReceber = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Numero do Documento] < " & mNumeroDocumento
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Documento] DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsDuplicataReceber.Close
    Set rsDuplicataReceber = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal nNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL(" WHERE Nome = " & preparaTexto(nNome), "")
    LocalizarNome = Localizar(1)
    If LocalizarNome Then
        AtribuiValor
    End If
    rsDuplicataReceber.Close
    Set rsDuplicataReceber = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Documento] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsDuplicataReceber.Close
    Set rsDuplicataReceber = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Numero do Documento] > " & mNumeroDocumento
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Documento] ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsDuplicataReceber.Close
    Set rsDuplicataReceber = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Documento] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsDuplicataReceber.Close
    Set rsDuplicataReceber = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Duplicata_Receber ( Empresa, [Numero do Documento], [Codigo do Cliente], [Data de Emissao], [Data do Periodo Inicial], "
    gSQL = gSQL & "[Data do Periodo Final], [Numero da Duplicata], [Data do Vencimento], [Valor do Custo Bancario], [Valor do Vencimento], "
    gSQL = gSQL & "[Numero do Movimento do Caixa], [Numero da Conta Bancaria] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mNumeroDocumento, ", ")
    Call sqlNumero(1, mCodigoCliente, ", ")
    Call sqlData(1, mDataEmissao, ", ")
    Call sqlData(1, mDataPeriodoInicial, ", ")
    Call sqlData(1, mDataPeriodoFinal, ", ")
    Call sqlTexto(1, mNumeroDuplicata, ", ")
    Call sqlData(1, mDataVencimento, ", ")
    Call sqlValor(1, mValorCustoBancario, ", ")
    Call sqlValor(1, mValorVencimento, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixa, ", ")
    Call sqlTexto(1, mNumeroContaBancaria, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nNumeroDocumento As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Duplicata_Receber SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", [Numero do Documento] = ", mNumeroDocumento)
    Call sqlNumero(2, ", [Codigo do Cliente] = ", mCodigoCliente)
    Call sqlData(2, ", [Data de Emissao] = ", mDataEmissao)
    Call sqlData(2, ", [Data do Periodo Inicial] = ", mDataPeriodoInicial)
    Call sqlData(2, ", [Data do Periodo Final] = ", mDataPeriodoFinal)
    Call sqlTexto(2, ", [Numero da Duplicata] = ", mNumeroDuplicata)
    Call sqlData(2, ", [Data do Vencimento] = ", mDataVencimento)
    Call sqlValor(2, ", [Valor do Custo Bancario] = ", mValorCustoBancario)
    Call sqlValor(2, ", [Valor do Vencimento] = ", mValorVencimento)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa] = ", mNumeroMovimentoCaixa)
    Call sqlTexto(2, ", [Numero da Conta Bancaria] = ", mNumeroContaBancaria)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Numero do Documento] = " & nNumeroDocumento
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nNumeroDocumento As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Duplicata_Receber"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Numero do Documento] = " & nNumeroDocumento
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Numero do Documento], [Codigo do Cliente], [Data de Emissao], [Data do Periodo Inicial], [Data do Periodo Final], [Numero da Duplicata], [Data do Vencimento], [Valor do Custo Bancario], [Valor do Vencimento], [Numero do Movimento do Caixa], [Numero da Conta Bancaria] FROM Duplicata_Receber" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsDuplicataReceber = New adodb.Recordset
    rsDuplicataReceber.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsDuplicataReceber.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsDuplicataReceber.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsDuplicataReceber("Empresa").Value
    NumeroDocumento = rsDuplicataReceber("Numero do Documento").Value
    CodigoCliente = rsDuplicataReceber("Codigo do Cliente").Value
    DataEmissao = rsDuplicataReceber("Data de Emissao").Value
    DataPeriodoInicial = rsDuplicataReceber("Data do Periodo Inicial").Value
    DataPeriodoFinal = rsDuplicataReceber("Data do Periodo Final").Value
    NumeroDuplicata = rsDuplicataReceber("Numero da Duplicata").Value
    DataVencimento = rsDuplicataReceber("Data do Vencimento").Value
    ValorCustoBancario = rsDuplicataReceber("Valor do Custo Bancario").Value
    ValorVencimento = rsDuplicataReceber("Valor do Vencimento").Value
    NumeroMovimentoCaixa = rsDuplicataReceber("Numero do Movimento do Caixa").Value
    NumeroContaBancaria = rsDuplicataReceber("Numero da Conta Bancaria").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function


Public Function ProximoNumeroDocumento(ByVal pEmpresa As Integer) As Long

On Error GoTo Err_rsDuplicataReceber

    ProximoNumeroDocumento = 1
    gSQL = ""
    gSQL = gSQL & "SELECT [Numero do Documento]"
    gSQL = gSQL & "  FROM Duplicata_Receber"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " ORDER BY [Numero do Documento] DESC"
    If Localizar(1) Then
        ProximoNumeroDocumento = rsDuplicataReceber("Numero do Documento").Value + 1
    End If
    rsDuplicataReceber.Close
    Set rsDuplicataReceber = Nothing
    
    gSQL = ""
    gSQL = gSQL & "SELECT [Numero do Documento]"
    gSQL = gSQL & "  FROM Baixa_Duplicata_Receber"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " ORDER BY [Numero do Documento] DESC"
    If Localizar(1) Then
        If ProximoNumeroDocumento <= rsDuplicataReceber("Numero do Documento").Value Then
            ProximoNumeroDocumento = rsDuplicataReceber("Numero do Documento").Value + 1
        End If
    End If
    rsDuplicataReceber.Close
    Set rsDuplicataReceber = Nothing
    Exit Function

Err_rsDuplicataReceber:
End Function


Public Function UltimoNumeroDuplicata(ByVal nEmpresa As Integer) As String

On Error GoTo trata_erro
    UltimoNumeroDuplicata = ""
    gSQL = "SELECT Convert(Int, [Numero da Duplicata]) AS Total"
    gSQL = gSQL & "  FROM Duplicata_Receber"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " ORDER BY Total DESC"
    If Localizar(0) Then
        If Not IsNull(rsDuplicataReceber("Total").Value) Then
            UltimoNumeroDuplicata = rsDuplicataReceber("Total").Value
        End If
    End If
    rsDuplicataReceber.Close
    Set rsDuplicataReceber = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function UltimoNumeroDuplicataConta(ByVal pEmpresa As Integer, ByVal pNumeroContaBancaria As String) As Long

On Error GoTo trata_erro
    
    UltimoNumeroDuplicataConta = 0
    gSQL = "SELECT Convert(Int, [Numero da Duplicata]) AS Numero"
    gSQL = gSQL & "  FROM Duplicata_Receber"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND [Numero da Conta Bancaria] = " & preparaTexto(pNumeroContaBancaria)
    gSQL = gSQL & " ORDER BY Numero DESC"
    If Localizar(1) Then
        If Not IsNull(rsDuplicataReceber("Numero").Value) Then
            UltimoNumeroDuplicataConta = rsDuplicataReceber("Numero").Value
        End If
    End If
    rsDuplicataReceber.Close
    Set rsDuplicataReceber = Nothing
    
    
    gSQL = "SELECT Convert(Int, [Numero da Duplicata]) AS Numero"
    gSQL = gSQL & "  FROM Baixa_Duplicata_Receber"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND [Numero da Conta Bancaria] = " & preparaTexto(pNumeroContaBancaria)
    gSQL = gSQL & " ORDER BY Numero DESC"
    If Localizar(1) Then
        If Not IsNull(rsDuplicataReceber("Numero").Value) Then
            If UltimoNumeroDuplicataConta < rsDuplicataReceber("Numero").Value Then
                UltimoNumeroDuplicataConta = rsDuplicataReceber("Numero").Value
            End If
        End If
    End If
    rsDuplicataReceber.Close
    Set rsDuplicataReceber = Nothing
    
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizaPrimeiroVencCliente(ByVal pCodigoCliente As Integer) As Boolean
    Dim xCondicao As String

On Error GoTo trata_erro

    LocalizaPrimeiroVencCliente = False
    xCondicao = " WHERE [Codigo do Cliente] = " & pCodigoCliente
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data do Vencimento] ASC")
    If Localizar(1) Then
        LocalizaPrimeiroVencCliente = True
        AtribuiValor
    End If
    rsDuplicataReceber.Close
    Set rsDuplicataReceber = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizaPrimeiroVencimento(ByVal pEmpresa As Integer) As Date

On Error GoTo trata_erro
    
    LocalizaPrimeiroVencimento = Date
    gSQL = "SELECT TOP 1 [Data do Vencimento] FROM Duplicata_Receber"
    If pEmpresa > 0 Then
        gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    End If
    gSQL = gSQL & " ORDER BY [Data do Vencimento] ASC"
    If Localizar(0) Then
        If Not IsNull(rsDuplicataReceber("Data do Vencimento").Value) Then
            LocalizaPrimeiroVencimento = rsDuplicataReceber("Data do Vencimento").Value
        End If
    End If
    rsDuplicataReceber.Close
    Set rsDuplicataReceber = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizaUltimoVencimento(ByVal pEmpresa As Integer) As Date

On Error GoTo trata_erro
    
    LocalizaUltimoVencimento = Date
    gSQL = "SELECT TOP 1 [Data do Vencimento] FROM Duplicata_Receber"
    If pEmpresa > 0 Then
        gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    End If
    gSQL = gSQL & " ORDER BY [Data do Vencimento] DESC"
    If Localizar(0) Then
        If Not IsNull(rsDuplicataReceber("Data do Vencimento").Value) Then
            LocalizaUltimoVencimento = rsDuplicataReceber("Data do Vencimento").Value
        End If
    End If
    rsDuplicataReceber.Close
    Set rsDuplicataReceber = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function



Public Function TotalEntreDatas(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pDataVencida As Date, ByVal pVencida As Boolean) As Currency

On Error GoTo trata_erro
    
    TotalEntreDatas = 0
    gSQL = "SELECT SUM([Valor do Vencimento]) as Total FROM Duplicata_Receber"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Data do Vencimento] >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND [Data do Vencimento] <= " & preparaData(pDataFinal)
    If pVencida Then
        gSQL = gSQL & " AND [Data do Vencimento] <= " & preparaData(pDataVencida)
    Else
        gSQL = gSQL & " AND [Data do Vencimento] > " & preparaData(pDataVencida)
    End If
    If Localizar(0) Then
        If Not IsNull(rsDuplicataReceber("Total").Value) Then
            TotalEntreDatas = rsDuplicataReceber("Total").Value
        End If
    End If
    rsDuplicataReceber.Close
    Set rsDuplicataReceber = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

Public Function TotalCliente(ByVal pEmpresa As Integer, ByVal pCliente As Integer) As Currency

On Error GoTo trata_erro
    
    TotalCliente = 0
    gSQL = "SELECT SUM([Valor do Vencimento]) as Total FROM Duplicata_Receber"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo do Cliente] = " & pCliente
    'gSQL = gSQL & " AND [Data do Vencimento] <= " & preparaData(pDataFinal)
'    If pVencida Then
'        gSQL = gSQL & " AND [Data do Vencimento] <= " & preparaData(pDataVencida)
'    Else
'        gSQL = gSQL & " AND [Data do Vencimento] > " & preparaData(pDataVencida)
'    End If
    If Localizar(0) Then
        If Not IsNull(rsDuplicataReceber("Total").Value) Then
            TotalCliente = rsDuplicataReceber("Total").Value
        End If
    End If
    rsDuplicataReceber.Close
    Set rsDuplicataReceber = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

