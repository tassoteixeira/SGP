VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBaixaDuplicataReceber"
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
Private mDataPagamento As Date
Private mValorJuros As Currency
Private mValorPago As Currency
Private mValorPagoChequeVista As Currency
Private mValorPagoChequePrazo As Currency
Private mPeriodo As Integer
Private mNumeroMovimentoCaixa As Long
Private mNumeroMovimentoCaixaBaixa As Long
Private mValorDesconto As Currency
Private mValorRestante As Currency
Private mDataVencimentoRestante As Date
Private mNumeroContaBancaria As String
Private mValorPagoBanco As Currency


Private rsBaixaDuplicataReceber As New adodb.Recordset


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


Public Property Get DataPagamento() As Date
    DataPagamento = mDataPagamento
End Property
Public Property Let DataPagamento(ByVal Valor As Date)
    mDataPagamento = Valor
End Property


Public Property Get ValorJuros() As Currency
    ValorJuros = mValorJuros
End Property
Public Property Let ValorJuros(ByVal Valor As Currency)
    mValorJuros = Valor
End Property


Public Property Get ValorPago() As Currency
    ValorPago = mValorPago
End Property
Public Property Let ValorPago(ByVal Valor As Currency)
    mValorPago = Valor
End Property


Public Property Get ValorPagoChequeVista() As Currency
    ValorPagoChequeVista = mValorPagoChequeVista
End Property
Public Property Let ValorPagoChequeVista(ByVal Valor As Currency)
    mValorPagoChequeVista = Valor
End Property


Public Property Get ValorPagoChequePrazo() As Currency
    ValorPagoChequePrazo = mValorPagoChequePrazo
End Property
Public Property Let ValorPagoChequePrazo(ByVal Valor As Currency)
    mValorPagoChequePrazo = Valor
End Property


Public Property Get Periodo() As Integer
    Periodo = mPeriodo
End Property
Public Property Let Periodo(ByVal Valor As Integer)
    mPeriodo = Valor
End Property


Public Property Get NumeroMovimentoCaixa() As Long
    NumeroMovimentoCaixa = mNumeroMovimentoCaixa
End Property
Public Property Let NumeroMovimentoCaixa(ByVal Valor As Long)
    mNumeroMovimentoCaixa = Valor
End Property


Public Property Get NumeroMovimentoCaixaBaixa() As Long
    NumeroMovimentoCaixaBaixa = mNumeroMovimentoCaixaBaixa
End Property
Public Property Let NumeroMovimentoCaixaBaixa(ByVal Valor As Long)
    mNumeroMovimentoCaixaBaixa = Valor
End Property


Public Property Get ValorDesconto() As Currency
    ValorDesconto = mValorDesconto
End Property
Public Property Let ValorDesconto(ByVal Valor As Currency)
    mValorDesconto = Valor
End Property


Public Property Get ValorRestante() As Currency
    ValorRestante = mValorRestante
End Property
Public Property Let ValorRestante(ByVal Valor As Currency)
    mValorRestante = Valor
End Property


Public Property Get DataVencimentoRestante() As Date
    DataVencimentoRestante = mDataVencimentoRestante
End Property
Public Property Let DataVencimentoRestante(ByVal Valor As Date)
    mDataVencimentoRestante = Valor
End Property


Public Property Get NumeroContaBancaria() As String
    NumeroContaBancaria = mNumeroContaBancaria
End Property
Public Property Let NumeroContaBancaria(ByVal Valor As String)
    mNumeroContaBancaria = Valor
End Property

Public Property Get ValorPagoBanco() As Currency
    ValorPagoBanco = mValorPagoBanco
End Property
Public Property Let ValorPagoBanco(ByVal Valor As Currency)
    mValorPagoBanco = Valor
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
    rsBaixaDuplicataReceber.Close
    Set rsBaixaDuplicataReceber = Nothing
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
    rsBaixaDuplicataReceber.Close
    Set rsBaixaDuplicataReceber = Nothing
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
    rsBaixaDuplicataReceber.Close
    Set rsBaixaDuplicataReceber = Nothing
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
    rsBaixaDuplicataReceber.Close
    Set rsBaixaDuplicataReceber = Nothing
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
    rsBaixaDuplicataReceber.Close
    Set rsBaixaDuplicataReceber = Nothing
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
    rsBaixaDuplicataReceber.Close
    Set rsBaixaDuplicataReceber = Nothing
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
    rsBaixaDuplicataReceber.Close
    Set rsBaixaDuplicataReceber = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Baixa_Duplicata_Receber ( Empresa, [Numero do Documento], [Codigo do Cliente], [Data de Emissao], [Data do Periodo Inicial], "
    gSQL = gSQL & "[Data do Periodo Final], [Numero da Duplicata], [Data do Vencimento], [Valor do Custo Bancario], [Valor do Vencimento], "
    gSQL = gSQL & "[Data do Pagamento], [Valor do Juros], [Valor Pago], [Valor Pago Cheque Vista], [Valor Pago Cheque Prazo], "
    gSQL = gSQL & "Periodo, [Numero do Movimento do Caixa], [Numero do Movimento do Caixa Baixa], [Valor do Desconto], "
    gSQL = gSQL & "[Valor Restante], [Data de Vencimento do Restante], [Numero da Conta Bancaria], [Valor Pago Banco] ) VALUES ( "
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
    Call sqlData(1, mDataPagamento, ", ")
    Call sqlValor(1, mValorJuros, ", ")
    Call sqlValor(1, mValorPago, ", ")
    Call sqlValor(1, mValorPagoChequeVista, ", ")
    Call sqlValor(1, mValorPagoChequePrazo, ", ")
    Call sqlNumero(1, mPeriodo, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixa, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixaBaixa, ", ")
    Call sqlValor(1, mValorDesconto, ", ")
    Call sqlValor(1, mValorRestante, ", ")
    Call sqlData(1, mDataVencimentoRestante, ", ")
    Call sqlTexto(1, mNumeroContaBancaria, ", ")
    Call sqlValor(1, mValorPagoBanco, ") ")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    Else
        Call CriaLogCadastroDll2("Classe:cBaixaDuplicataReceber - Método:Incluir", "Não foi possível incluir registro.", gSQL)
    End If
    Exit Function

Err_Incluir:
    Call CriaLogCadastroDll2("Classe:cBaixaDuplicataReceber - Método:Incluir", Err.Description, gSQL)
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nNumeroDocumento As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Baixa_Duplicata_Receber SET "
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
    Call sqlData(2, ", [Data do Pagamento] = ", mDataPagamento)
    Call sqlValor(2, ", [Valor do Juros] = ", mValorJuros)
    Call sqlValor(2, ", [Valor Pago] = ", mValorPago)
    Call sqlValor(2, ", [Valor Pago Cheque Vista] = ", mValorPagoChequeVista)
    Call sqlValor(2, ", [Valor Pago Cheque Prazo] = ", mValorPagoChequePrazo)
    Call sqlNumero(2, ", Periodo = ", mPeriodo)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa] = ", mNumeroMovimentoCaixa)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa Baixa] = ", mNumeroMovimentoCaixaBaixa)
    Call sqlValor(2, ", [Valor do Desconto] = ", mValorDesconto)
    Call sqlValor(2, ", [Valor Restante] = ", mValorRestante)
    Call sqlData(2, ", [Data de Vencimento do Restante] = ", mDataVencimentoRestante)
    Call sqlTexto(2, ", [Numero da Conta Bancaria] = ", mNumeroContaBancaria)
    Call sqlValor(2, ", [Valor Pago Banco] = ", mValorPagoBanco)
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
    gSQL = "DELETE FROM Baixa_Duplicata_Receber"
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
    PreparaSQL = Trim("SELECT Empresa, [Numero do Documento], [Codigo do Cliente], [Data de Emissao], [Data do Periodo Inicial], [Data do Periodo Final], [Numero da Duplicata], [Data do Vencimento], [Valor do Custo Bancario], [Valor do Vencimento], [Data do Pagamento], [Valor do Juros], [Valor Pago], [Valor Pago Cheque Vista], [Valor Pago Cheque Prazo], Periodo, [Numero do Movimento do Caixa], [Numero do Movimento do Caixa Baixa], [Valor do Desconto], [Valor Restante], [Data de Vencimento do Restante], [Numero da Conta Bancaria], [Valor Pago Banco] FROM Baixa_Duplicata_Receber" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsBaixaDuplicataReceber = New adodb.Recordset
    rsBaixaDuplicataReceber.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsBaixaDuplicataReceber.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsBaixaDuplicataReceber.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsBaixaDuplicataReceber("Empresa").Value
    NumeroDocumento = rsBaixaDuplicataReceber("Numero do Documento").Value
    CodigoCliente = rsBaixaDuplicataReceber("Codigo do Cliente").Value
    DataEmissao = rsBaixaDuplicataReceber("Data de Emissao").Value
    DataPeriodoInicial = rsBaixaDuplicataReceber("Data do Periodo Inicial").Value
    DataPeriodoFinal = rsBaixaDuplicataReceber("Data do Periodo Final").Value
    NumeroDuplicata = rsBaixaDuplicataReceber("Numero da Duplicata").Value
    DataVencimento = rsBaixaDuplicataReceber("Data do Vencimento").Value
    ValorCustoBancario = rsBaixaDuplicataReceber("Valor do Custo Bancario").Value
    ValorVencimento = rsBaixaDuplicataReceber("Valor do Vencimento").Value
    DataPagamento = rsBaixaDuplicataReceber("Data do Pagamento").Value
    ValorJuros = rsBaixaDuplicataReceber("Valor do Juros").Value
    ValorPago = rsBaixaDuplicataReceber("Valor Pago").Value
    ValorPagoChequeVista = rsBaixaDuplicataReceber("Valor Pago Cheque Vista").Value
    ValorPagoChequePrazo = rsBaixaDuplicataReceber("Valor Pago Cheque Prazo").Value
    Periodo = rsBaixaDuplicataReceber("Periodo").Value
    NumeroMovimentoCaixa = rsBaixaDuplicataReceber("Numero do Movimento do Caixa").Value
    NumeroMovimentoCaixaBaixa = rsBaixaDuplicataReceber("Numero do Movimento do Caixa Baixa").Value
    ValorDesconto = rsBaixaDuplicataReceber("Valor do Desconto").Value
    ValorRestante = rsBaixaDuplicataReceber("Valor Restante").Value
    If IsNull(rsBaixaDuplicataReceber("Data de Vencimento do Restante").Value) Then
        DataVencimentoRestante = "00:00:00"
    Else
        DataVencimentoRestante = rsBaixaDuplicataReceber("Data de Vencimento do Restante").Value
    End If
    NumeroContaBancaria = rsBaixaDuplicataReceber("Numero da Conta Bancaria").Value
    ValorPagoBanco = rsBaixaDuplicataReceber("Valor Pago Banco").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function


Public Function UltimoNumeroDuplicata(ByVal nEmpresa As Integer) As String

On Error GoTo trata_erro
    UltimoNumeroDuplicata = ""
    gSQL = "SELECT Convert(Int, [Numero da Duplicata]) AS Total"
    gSQL = gSQL & "  FROM Baixa_Duplicata_Receber"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " ORDER BY Total DESC"
    If Localizar(0) Then
        If Not IsNull(rsBaixaDuplicataReceber("Total").Value) Then
            UltimoNumeroDuplicata = rsBaixaDuplicataReceber("Total").Value
        End If
    End If
    rsBaixaDuplicataReceber.Close
    Set rsBaixaDuplicataReceber = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

