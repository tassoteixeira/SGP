VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoAfericao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mNomeTabela As String
Private mEmpresa As Integer
Private mData As Date
Private mPeriodo As Integer
Private mCodigoBomba As Integer
Private mQuantidade As Currency
Private mPrecoCusto As Currency
Private mPrecoVenda As Currency
Private mValorTotal As Currency
Private mTipoCombustivel As String
Private mTransferencia As Boolean
Private mNumeroMovimentoCaixaCredito As Long
Private mNumeroMovimentoCaixaDebito As Long
Private mNumeroIlha As Integer
Private mDataAbastecimento As Date
Private mHoraAbastecimento As Date


Private rsMovAfericao As New adodb.Recordset
'Private MovCaixa As cMovimentoCaixa


'Inicio das Propriedades da Classe
Public Property Get NomeTabela() As String
    NomeTabela = mNomeTabela
End Property
Public Property Let NomeTabela(ByVal Valor As String)
    mNomeTabela = Valor
End Property


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


Public Property Get CodigoBomba() As Integer
    CodigoBomba = mCodigoBomba
End Property
Public Property Let CodigoBomba(ByVal Valor As Integer)
    mCodigoBomba = Valor
End Property


Public Property Get Quantidade() As Currency
    Quantidade = mQuantidade
End Property
Public Property Let Quantidade(ByVal Valor As Currency)
    mQuantidade = Valor
End Property


Public Property Get PrecoCusto() As Currency
    PrecoCusto = mPrecoCusto
End Property
Public Property Let PrecoCusto(ByVal Valor As Currency)
    mPrecoCusto = Valor
End Property


Public Property Get PrecoVenda() As Currency
    PrecoVenda = mPrecoVenda
End Property
Public Property Let PrecoVenda(ByVal Valor As Currency)
    mPrecoVenda = Valor
End Property


Public Property Get ValorTotal() As Currency
    ValorTotal = mValorTotal
End Property
Public Property Let ValorTotal(ByVal Valor As Currency)
    mValorTotal = Valor
End Property


Public Property Get TipoCombustivel() As String
    TipoCombustivel = mTipoCombustivel
End Property
Public Property Let TipoCombustivel(ByVal Valor As String)
    mTipoCombustivel = Valor
End Property


Public Property Get Transferencia() As Boolean
    Transferencia = mTransferencia
End Property
Public Property Let Transferencia(ByVal Valor As Boolean)
    mTransferencia = Valor
End Property


Public Property Get NumeroMovimentoCaixaCredito() As Long
    NumeroMovimentoCaixaCredito = mNumeroMovimentoCaixaCredito
End Property
Public Property Let NumeroMovimentoCaixaCredito(ByVal Valor As Long)
    mNumeroMovimentoCaixaCredito = Valor
End Property


Public Property Get NumeroMovimentoCaixaDebito() As Long
    NumeroMovimentoCaixaDebito = mNumeroMovimentoCaixaDebito
End Property
Public Property Let NumeroMovimentoCaixaDebito(ByVal Valor As Long)
    mNumeroMovimentoCaixaDebito = Valor
End Property


Public Property Get NumeroIlha() As Integer
    NumeroIlha = mNumeroIlha
End Property
Public Property Let NumeroIlha(ByVal Valor As Integer)
    mNumeroIlha = Valor
End Property

Public Property Get DataAbastecimento() As Date
    DataAbastecimento = mDataAbastecimento
End Property
Public Property Let DataAbastecimento(ByVal Valor As Date)
    mDataAbastecimento = Valor
End Property

Public Property Get HoraAbastecimento() As Date
    HoraAbastecimento = mHoraAbastecimento
End Property
Public Property Let HoraAbastecimento(ByVal Valor As Date)
    mHoraAbastecimento = Valor
End Property


'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nPeriodo As Integer, ByVal nCodigoBomba As Integer, ByVal nTransferencia As Boolean) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(nData)
    xCondicao = xCondicao & " AND Periodo = " & nPeriodo
    xCondicao = xCondicao & " AND [Codigo da Bomba] = " & nCodigoBomba
    xCondicao = xCondicao & " AND Transferencia = " & CCur(nTransferencia)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsMovAfericao.Close
    Set rsMovAfericao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

Public Function LocalizarAnterior() As Boolean
Dim xSQL As String
Dim xOrdem As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarAnterior = False
    xRegIgual = False
    xSQL = " WHERE Empresa = " & mEmpresa
    xSQL = xSQL & " AND Data <= " & preparaData(mData)
    xOrdem = "ORDER BY Empresa DESC, Data DESC, Periodo DESC, [Codigo da Bomba] DESC, Transferencia DESC"
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(55) Then
        If rsMovAfericao.RecordCount > 1 Then
            Do Until rsMovAfericao.EOF
                If xRegIgual = False Then
                    If rsMovAfericao("Data").Value = mData And rsMovAfericao("Periodo").Value = mPeriodo And rsMovAfericao("Codigo da Bomba").Value = mCodigoBomba And rsMovAfericao("Transferencia").Value = mTransferencia Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsMovAfericao.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsMovAfericao.Close
    Set rsMovAfericao = Nothing
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
    rsMovAfericao.Close
    Set rsMovAfericao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Data ASC, Periodo ASC, [Codigo da Bomba] ASC, Transferencia ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsMovAfericao.Close
    Set rsMovAfericao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xSQL As String
Dim xOrdem As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarProximo = False
    xRegIgual = False
    xSQL = " WHERE Empresa = " & mEmpresa
    xSQL = xSQL & " AND Data >= " & preparaData(mData)
    xOrdem = "ORDER BY Empresa ASC, Data ASC, Periodo ASC, [Codigo da Bomba] ASC, Transferencia ASC"
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(55) Then
        If rsMovAfericao.RecordCount > 1 Then
            Do Until rsMovAfericao.EOF
                If xRegIgual = False Then
                    If rsMovAfericao("Data").Value = mData And rsMovAfericao("Periodo").Value = mPeriodo And rsMovAfericao("Codigo da Bomba").Value = mCodigoBomba And rsMovAfericao("Transferencia").Value = mTransferencia Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsMovAfericao.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMovAfericao.Close
    Set rsMovAfericao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Data DESC, Periodo DESC, [Codigo da Bomba] DESC, Transferencia DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsMovAfericao.Close
    Set rsMovAfericao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO " & mNomeTabela & " ( Empresa, Data, Periodo, [Codigo da Bomba], Quantidade, "
    gSQL = gSQL & "[Preco de Custo], [Preco de Venda], [Valor Total], [Tipo de Combustivel], "
    gSQL = gSQL & "Transferencia, [Numero do Movimento do Caixa Credito], [Numero do Movimento do Caixa Debito], "
    gSQL = gSQL & "[Numero da Ilha]) VALUES ( "
    'gSQL = gSQL & "[Numero da Ilha], [Data do Abastecimento], [Hora do Abastecimento] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlNumero(1, mPeriodo, ", ")
    Call sqlNumero(1, mCodigoBomba, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlValor(1, mPrecoCusto, ", ")
    Call sqlValor(1, mPrecoVenda, ", ")
    Call sqlValor(1, mValorTotal, ", ")
    Call sqlTexto(1, mTipoCombustivel, ", ")
    Call sqlBoolean(1, mTransferencia, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixaCredito, ", ")
    Call sqlNumero(1, mNumeroMovimentoCaixaDebito, ", ")
    Call sqlNumero(1, mNumeroIlha, " )")
'    Call sqlData(1, mDataAbastecimento, ", ")
'    Call sqlHora(1, mHoraAbastecimento, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nPeriodo As Integer, ByVal nCodigoBomba As Integer, ByVal nTransferencia As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE " & mNomeTabela & " SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlNumero(2, ", Periodo = ", mPeriodo)
    Call sqlNumero(2, ", [Codigo da Bomba] = ", mCodigoBomba)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlValor(2, ", [Preco de Custo] = ", mPrecoCusto)
    Call sqlValor(2, ", [Preco de Venda] = ", mPrecoVenda)
    Call sqlValor(2, ", [Valor Total] = ", mValorTotal)
    Call sqlTexto(2, ", [Tipo de Combustivel] = ", mTipoCombustivel)
    Call sqlBoolean(2, ", Transferencia = ", mTransferencia)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa Credito] = ", mNumeroMovimentoCaixaCredito)
    Call sqlNumero(2, ", [Numero do Movimento do Caixa Debito] = ", mNumeroMovimentoCaixaDebito)
    Call sqlNumero(2, ", [Numero da Ilha] = ", mNumeroIlha)
'    Call sqlData(2, ", [Data do Abastecimento] = ", mDataAbastecimento)
'    Call sqlHora(2, ", [Hora do Abastecimento] = ", mHoraAbastecimento)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND Periodo = " & nPeriodo
    gSQL = gSQL & " AND [Codigo da Bomba] = " & nCodigoBomba
    gSQL = gSQL & " AND Transferencia = " & preparaBooleano(nTransferencia)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function AlteraPrecoCusto(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nTipoCombustivel As String, ByVal nPrecoCusto As Currency) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_AlteraPrecoCusto

    AlteraPrecoCusto = False
    gSQL = "UPDATE " & mNomeTabela & " SET "
    Call sqlValor(2, "[Preco de Custo] = ", nPrecoCusto)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto(nTipoCombustivel)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlteraPrecoCusto = True
    End If
    Exit Function

Err_AlteraPrecoCusto:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nPeriodo As Integer, ByVal nCodigoBomba As Integer, ByVal nTransferencia As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    If LocalizarCodigo(nEmpresa, nData, nPeriodo, nCodigoBomba, nTransferencia) Then
    End If
    gSQL = "DELETE FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND Periodo = " & nPeriodo
    gSQL = gSQL & " AND [Codigo da Bomba] = " & nCodigoBomba
    gSQL = gSQL & " AND Transferencia = " & preparaBooleano(nTransferencia)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


Public Function TotalPeriodo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pTransferencia As Boolean) As Currency

On Error GoTo trata_erro
    TotalPeriodo = 0
    gSQL = "SELECT SUM([Valor Total]) as Total FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    If pPeriodo > 0 Then
        gSQL = gSQL & " AND Periodo = " & pPeriodo
    End If
    If pTransferencia = False Then
        gSQL = gSQL & " AND Transferencia = 0"
    Else
        gSQL = gSQL & " AND Transferencia = -1"
    End If
    If Localizar(0) Then
        If Not IsNull(rsMovAfericao("Total").Value) Then
            TotalPeriodo = rsMovAfericao("Total").Value
        End If
    End If
    rsMovAfericao.Close
    Set rsMovAfericao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
'Fim Métodos da Classe


Public Function TotalQtdPeriodoCombustivel(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pTipoCombustivel As String, ByVal pDataPeriodo As String) As Currency

On Error GoTo trata_erro
    TotalQtdPeriodoCombustivel = 0
    gSQL = "SELECT SUM(Quantidade) as Total FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND Data <= " & preparaData(pDataFinal)
    gSQL = gSQL & " AND Periodo >= " & pPeriodoInicial
    gSQL = gSQL & " AND Periodo <= " & pPeriodoFinal
    If pTipoCombustivel <> "" Then
        gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    End If
    If pDataPeriodo <> "" Then
        gSQL = gSQL & "    AND CONVERT(VARCHAR, Data, 103) + CONVERT(VARCHAR, Periodo) IN " & pDataPeriodo
    End If
    If Localizar(0) Then
        If Not IsNull(rsMovAfericao("Total").Value) Then
            TotalQtdPeriodoCombustivel = rsMovAfericao("Total").Value
        End If
    End If
    rsMovAfericao.Close
    Set rsMovAfericao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TransfereDadosLMC(ByVal pEmpresa As Integer, ByVal pSomenteVerifica As Boolean) As Boolean
    Dim xData As Date
    Dim rsAuxiliar As New adodb.Recordset

On Error GoTo trata_erro

    TransfereDadosLMC = False
    
    'Busca Ultima Data do Movimento LMC
    xData = UltimaData(pEmpresa)
    If xData = "00:00:00" Then
        Exit Function
    End If
    
    'Busca Movimento de Afericao para Transferencia
    xData = xData + 1
    gSQL = ""
    gSQL = gSQL & "   SELECT Empresa, Data, Periodo, [Codigo da Bomba], Quantidade, [Preco de Custo],"
    gSQL = gSQL & "          [Preco de Venda], [Valor Total], [Tipo de Combustivel], Transferencia,"
    gSQL = gSQL & "          [Numero do Movimento do Caixa Credito], [Numero do Movimento do Caixa Debito],"
    gSQL = gSQL & "          [Numero da Ilha]"
    'gSQL = gSQL & "          [Numero da Ilha], [Data do Abastecimento], [Hora do Abastecimento]"
    gSQL = gSQL & "     FROM Movimento_Afericao"
    gSQL = gSQL & "    WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "      AND Data >= " & preparaData(xData)
    gSQL = gSQL & " ORDER BY Data, Periodo"
    Set rsAuxiliar = New adodb.Recordset
    rsAuxiliar.CursorLocation = adUseClient
    rsAuxiliar.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    
    'Grava Afericoes
    If Not rsAuxiliar.EOF Then
        If pSomenteVerifica Then
            TransfereDadosLMC = True
            Exit Function
        End If
        Do Until rsAuxiliar.EOF
            Empresa = rsAuxiliar("Empresa").Value
            Data = rsAuxiliar("Data").Value
            Periodo = rsAuxiliar("Periodo").Value
            CodigoBomba = rsAuxiliar("Codigo da Bomba").Value
            Quantidade = rsAuxiliar("Quantidade").Value
            PrecoCusto = rsAuxiliar("Preco de Custo").Value
            PrecoVenda = rsAuxiliar("Preco de Venda").Value
            ValorTotal = rsAuxiliar("Valor Total").Value
            TipoCombustivel = rsAuxiliar("Tipo de Combustivel").Value
            Transferencia = rsAuxiliar("Transferencia").Value
            NumeroMovimentoCaixaCredito = rsAuxiliar("Numero do Movimento do Caixa Credito").Value
            NumeroMovimentoCaixaDebito = rsAuxiliar("Numero do Movimento do Caixa Debito").Value
            NumeroIlha = rsAuxiliar("Numero da Ilha").Value
            
            'Serão removidos da classe
            DataAbastecimento = CDate("00:00:00")
            HoraAbastecimento = CDate("00:00:00")
            
            If Not Incluir Then
                MsgBox "Registro de Aferição não foi gravado!", vbInformation, "Erro de Integridade"
            End If
            rsAuxiliar.MoveNext
        Loop
        TransfereDadosLMC = True
    End If
    
    rsAuxiliar.Close
    Set rsAuxiliar = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function UltimaData(ByVal pEmpresa As Integer) As Date

On Error GoTo trata_erro

    UltimaData = "01/01/1900"
    gSQL = "SELECT Data FROM Movimento_Afericao_LMC WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " ORDER BY Data DESC, Periodo DESC, [Codigo da Bomba] DESC, Transferencia DESC"
    If Localizar(1) Then
        UltimaData = rsMovAfericao("Data").Value
    End If
    rsMovAfericao.Close
    Set rsMovAfericao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ValorTotalCustoPeriodoCombustivel(ByVal nEmpresa As Integer, ByVal nDataInicial As Date, ByVal nDataFinal As Date, ByVal nPeriodoInicial As Integer, ByVal nPeriodoFinal As Integer, ByVal nTipoCombustivel As String, ByVal pDataPeriodo As String) As Currency

On Error GoTo trata_erro
    ValorTotalCustoPeriodoCombustivel = 0
    gSQL = "SELECT SUM([Preco de Custo] * Quantidade) as Total FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data >= " & preparaData(nDataInicial)
    gSQL = gSQL & " AND Data <= " & preparaData(nDataFinal)
    gSQL = gSQL & " AND Periodo >= " & nPeriodoInicial
    gSQL = gSQL & " AND Periodo <= " & nPeriodoFinal
    gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto(nTipoCombustivel)
    If pDataPeriodo <> "" Then
        gSQL = gSQL & "    AND CONVERT(VARCHAR, Data, 103) + CONVERT(VARCHAR, Periodo) IN " & pDataPeriodo
    End If
    If Localizar(0) Then
        If Not IsNull(rsMovAfericao("Total").Value) Then
            ValorTotalCustoPeriodoCombustivel = rsMovAfericao("Total").Value
        End If
    End If
    rsMovAfericao.Close
    Set rsMovAfericao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ValorTotalPeriodoCombustivel(ByVal nEmpresa As Integer, ByVal nDataInicial As Date, ByVal nDataFinal As Date, ByVal nPeriodoInicial As Integer, ByVal nPeriodoFinal As Integer, ByVal nTipoCombustivel As String, ByVal pDataPeriodo As String) As Currency

On Error GoTo trata_erro
    ValorTotalPeriodoCombustivel = 0
    gSQL = "SELECT SUM([Valor Total]) as Total FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data >= " & preparaData(nDataInicial)
    gSQL = gSQL & " AND Data <= " & preparaData(nDataFinal)
    gSQL = gSQL & " AND Periodo >= " & nPeriodoInicial
    gSQL = gSQL & " AND Periodo <= " & nPeriodoFinal
    gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto(nTipoCombustivel)
    If pDataPeriodo <> "" Then
        gSQL = gSQL & "    AND CONVERT(VARCHAR, Data, 103) + CONVERT(VARCHAR, Periodo) IN " & pDataPeriodo
    End If
    If Localizar(0) Then
        If Not IsNull(rsMovAfericao("Total").Value) Then
            ValorTotalPeriodoCombustivel = rsMovAfericao("Total").Value
        End If
    End If
    rsMovAfericao.Close
    Set rsMovAfericao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalQtdPeriodoBico(ByVal nEmpresa As Integer, ByVal nDataInicial As Date, ByVal nDataFinal As Date, ByVal nPeriodoInicial As Integer, ByVal nPeriodoFinal As Integer, ByVal nBico As Integer) As Currency

On Error GoTo trata_erro
    TotalQtdPeriodoBico = 0
    gSQL = "SELECT SUM(Quantidade) AS Total FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data >= " & preparaData(nDataInicial)
    gSQL = gSQL & " AND Data <= " & preparaData(nDataFinal)
    gSQL = gSQL & " AND Periodo >= " & nPeriodoInicial
    gSQL = gSQL & " AND Periodo <= " & nPeriodoFinal
    gSQL = gSQL & " AND [Codigo da Bomba] = " & nBico
    If Localizar(0) Then
        If Not IsNull(rsMovAfericao("Total").Value) Then
            TotalQtdPeriodoBico = rsMovAfericao("Total").Value
        End If
    End If
    rsMovAfericao.Close
    Set rsMovAfericao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalQtdBico_DataPeriodo(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pBico As Integer, ByVal pPeriodoInicial As Integer, ByVal pPeriodoFinal As Integer, ByVal pDataPeriodo As String) As Currency

On Error GoTo trata_erro
    TotalQtdBico_DataPeriodo = 0
    gSQL = "SELECT SUM(Quantidade) AS Total FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data >= " & preparaData(pDataInicial)
    gSQL = gSQL & " AND Data <= " & preparaData(pDataFinal)
    gSQL = gSQL & " AND Periodo >= " & pPeriodoInicial
    gSQL = gSQL & " AND Periodo <= " & pPeriodoFinal
    gSQL = gSQL & " AND [Codigo da Bomba] = " & pBico
    If pDataPeriodo <> "" Then
        gSQL = gSQL & "    AND CONVERT(VARCHAR, Data, 103) + CONVERT(VARCHAR, Periodo) IN " & pDataPeriodo
    End If
    If Localizar(0) Then
        If Not IsNull(rsMovAfericao("Total").Value) Then
            TotalQtdBico_DataPeriodo = rsMovAfericao("Total").Value
        End If
    End If
    rsMovAfericao.Close
    Set rsMovAfericao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ValorTotalPeriodoBico(ByVal nEmpresa As Integer, ByVal nDataInicial As Date, ByVal nDataFinal As Date, ByVal nPeriodoInicial As Integer, ByVal nPeriodoFinal As Integer, ByVal nBico As Integer) As Currency

On Error GoTo trata_erro
    ValorTotalPeriodoBico = 0
    gSQL = "SELECT SUM([Valor Total]) AS Total FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data >= " & preparaData(nDataInicial)
    gSQL = gSQL & " AND Data <= " & preparaData(nDataFinal)
    gSQL = gSQL & " AND Periodo >= " & nPeriodoInicial
    gSQL = gSQL & " AND Periodo <= " & nPeriodoFinal
    gSQL = gSQL & " AND [Codigo da Bomba] = " & nBico
    If Localizar(0) Then
        If Not IsNull(rsMovAfericao("Total").Value) Then
            ValorTotalPeriodoBico = rsMovAfericao("Total").Value
        End If
    End If
    rsMovAfericao.Close
    Set rsMovAfericao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
'Fim Métodos da Classe

'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    Dim xString As String
    xString = ""
    xString = xString & "Empresa, Data, Periodo, [Codigo da Bomba], Quantidade, [Preco de Custo], [Preco de Venda], "
    xString = xString & "[Valor Total], [Tipo de Combustivel], Transferencia, [Numero do Movimento do Caixa Credito], "
    xString = xString & "[Numero do Movimento do Caixa Debito], [Numero da Ilha]"
    'xString = xString & "[Numero do Movimento do Caixa Debito], [Numero da Ilha], [Data do Abastecimento], [Hora do Abastecimento]"
    PreparaSQL = Trim("SELECT " & xString & " FROM " & mNomeTabela & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovAfericao = New adodb.Recordset
    rsMovAfericao.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovAfericao.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovAfericao.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovAfericao("Empresa").Value
    Data = rsMovAfericao("Data").Value
    Periodo = rsMovAfericao("Periodo").Value
    CodigoBomba = rsMovAfericao("Codigo da Bomba").Value
    Quantidade = rsMovAfericao("Quantidade").Value
    PrecoCusto = rsMovAfericao("Preco de Custo").Value
    PrecoVenda = rsMovAfericao("Preco de Venda").Value
    ValorTotal = rsMovAfericao("Valor Total").Value
    TipoCombustivel = rsMovAfericao("Tipo de Combustivel").Value
    Transferencia = rsMovAfericao("Transferencia").Value
    NumeroMovimentoCaixaCredito = rsMovAfericao("Numero do Movimento do Caixa Credito").Value
    NumeroMovimentoCaixaDebito = rsMovAfericao("Numero do Movimento do Caixa Debito").Value
    NumeroIlha = rsMovAfericao("Numero da Ilha").Value
   
    'serão removidos da classe
    DataAbastecimento = CDate("00:00:00")
    HoraAbastecimento = CDate("00:00:00")
        
    
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function


'Private Sub Class_Initialize()
'    Set MovCaixa = New cMovimentoCaixa
'End Sub


'Private Sub Class_Terminate()
'    Set MovCaixa = Nothing
'End Sub
