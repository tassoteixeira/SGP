VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoAbastecimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mBico As Integer
Private mData As Date
Private mHora As Date
Private mTempoAbastecimento As String
Private mCodigoProduto As Integer
Private mValorUnitario As Currency
Private mQuantidade As Currency
Private mValorTotal As Currency
Private mAcerto As Boolean
Private mNumeroCupom As Long
Private mCodigoECF As Integer
Private mNumeroAbastecimento As String
Private mEncerrante As Currency
Private mStringAutomacao As String
Private mDocumentoGerado As String
Private mComplementoDocumentoGerado As String
Private mValorDesconto As Currency
Private mEncerranteInicial As Currency
Private mTipoCombustivel As String
Private mCodigoFuncionario As Integer
Private mPeriodo As Integer
Private mValorDescontoNivelPreco As Currency
Private mValorAcrescimoNivelPreco As Currency

Private rsMovAbastecimento As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get Bico() As Integer
    Bico = mBico
End Property
Public Property Let Bico(ByVal Valor As Integer)
    mBico = Valor
End Property


Public Property Get Data() As Date
    Data = mData
End Property
Public Property Let Data(ByVal Valor As Date)
    mData = Valor
End Property


Public Property Get Hora() As Date
    Hora = mHora
End Property
Public Property Let Hora(ByVal Valor As Date)
    mHora = Valor
End Property


Public Property Get TempoAbastecimento() As String
    TempoAbastecimento = mTempoAbastecimento
End Property
Public Property Let TempoAbastecimento(ByVal Valor As String)
    mTempoAbastecimento = Valor
End Property


Public Property Get CodigoProduto() As Integer
    CodigoProduto = mCodigoProduto
End Property
Public Property Let CodigoProduto(ByVal Valor As Integer)
    mCodigoProduto = Valor
End Property


Public Property Get ValorUnitario() As Currency
    ValorUnitario = mValorUnitario
End Property
Public Property Let ValorUnitario(ByVal Valor As Currency)
    mValorUnitario = Valor
End Property


Public Property Get Quantidade() As Currency
    Quantidade = mQuantidade
End Property
Public Property Let Quantidade(ByVal Valor As Currency)
    mQuantidade = Valor
End Property


Public Property Get ValorTotal() As Currency
    ValorTotal = mValorTotal
End Property
Public Property Let ValorTotal(ByVal Valor As Currency)
    mValorTotal = Valor
End Property


Public Property Get Acerto() As Boolean
    Acerto = mAcerto
End Property
Public Property Let Acerto(ByVal Valor As Boolean)
    mAcerto = Valor
End Property


Public Property Get NumeroCupom() As Long
    NumeroCupom = mNumeroCupom
End Property
Public Property Let NumeroCupom(ByVal Valor As Long)
    mNumeroCupom = Valor
End Property


Public Property Get CodigoECF() As Integer
    CodigoECF = mCodigoECF
End Property
Public Property Let CodigoECF(ByVal Valor As Integer)
    mCodigoECF = Valor
End Property


Public Property Get NumeroAbastecimento() As String
    NumeroAbastecimento = mNumeroAbastecimento
End Property
Public Property Let NumeroAbastecimento(ByVal Valor As String)
    mNumeroAbastecimento = Valor
End Property


Public Property Get Encerrante() As Currency
    Encerrante = mEncerrante
End Property
Public Property Let Encerrante(ByVal Valor As Currency)
    mEncerrante = Valor
End Property


Public Property Get StringAutomacao() As String
    StringAutomacao = mStringAutomacao
End Property
Public Property Let StringAutomacao(ByVal Valor As String)
    mStringAutomacao = Valor
End Property


Public Property Get DocumentoGerado() As String
    DocumentoGerado = mDocumentoGerado
End Property
Public Property Let DocumentoGerado(ByVal Valor As String)
    mDocumentoGerado = Valor
End Property


Public Property Get ComplementoDocumentoGerado() As String
    ComplementoDocumentoGerado = mComplementoDocumentoGerado
End Property
Public Property Let ComplementoDocumentoGerado(ByVal Valor As String)
    mComplementoDocumentoGerado = Valor
End Property


Public Property Get ValorDesconto() As Currency
    ValorDesconto = mValorDesconto
End Property
Public Property Let ValorDesconto(ByVal Valor As Currency)
    mValorDesconto = Valor
End Property


Public Property Get EncerranteInicial() As Currency
    EncerranteInicial = mEncerranteInicial
End Property
Public Property Let EncerranteInicial(ByVal Valor As Currency)
    mEncerranteInicial = Valor
End Property


Public Property Get TipoCombustivel() As String
    TipoCombustivel = mTipoCombustivel
End Property
Public Property Let TipoCombustivel(ByVal Valor As String)
    mTipoCombustivel = Valor
End Property


Public Property Get CodigoFuncionario() As Integer
    CodigoFuncionario = mCodigoFuncionario
End Property
Public Property Let CodigoFuncionario(ByVal Valor As Integer)
    mCodigoFuncionario = Valor
End Property


Public Property Get Periodo() As Integer
    Periodo = mPeriodo
End Property
Public Property Let Periodo(ByVal Valor As Integer)
    mPeriodo = Valor
End Property

Public Property Get ValorDescontoNivelPreco() As Currency
    ValorDescontoNivelPreco = mValorDescontoNivelPreco
End Property
Public Property Let ValorDescontoNivelPreco(ByVal Valor As Currency)
    mValorDescontoNivelPreco = Valor
End Property
Public Property Get ValorAcrescimoNivelPreco() As Currency
    ValorAcrescimoNivelPreco = mValorAcrescimoNivelPreco
End Property
Public Property Let ValorAcrescimoNivelPreco(ByVal Valor As Currency)
    mValorAcrescimoNivelPreco = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pHora As Date, ByVal pBico As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND " & preparaHoraConsulta("Hora", "=", pHora)
    xCondicao = xCondicao & " AND Bico = " & pBico
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsMovAbastecimento.Close
    Set rsMovAbastecimento = Nothing
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
    xCondicao = xCondicao & " AND Hora < " & preparaHora(mHora)
    xCondicao = xCondicao & " AND Bico < " & mBico
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Data DESC, Hora DESC, Bico DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsMovAbastecimento.Close
    Set rsMovAbastecimento = Nothing
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
    xCondicao = xCondicao & " AND Hora > " & preparaHora(mHora)
    xCondicao = xCondicao & " AND Bico > " & mBico
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Data ASC, Hora ASC, Bico ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsMovAbastecimento.Close
    Set rsMovAbastecimento = Nothing
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
    xCondicao = xCondicao & " AND Hora > " & preparaHora(mHora)
    xCondicao = xCondicao & " AND Bico > " & mBico
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Data ASC, Hora ASC, Bico ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMovAbastecimento.Close
    Set rsMovAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ExisteStringAutomacao(ByVal pEmpresa As Integer, ByVal pStringAutomacao As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    ExisteStringAutomacao = False
    gSQL = ""
    gSQL = gSQL & "SELECT [String da Automacao]"
    gSQL = gSQL & "  FROM Movimento_Abastecimento"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND [String da Automacao] = " & preparaTexto(pStringAutomacao)
    ExisteStringAutomacao = Localizar(1)
    rsMovAbastecimento.Close
    Set rsMovAbastecimento = Nothing
    
    If ExisteStringAutomacao = False Then
        gSQL = ""
        gSQL = gSQL & "SELECT [String da Automacao]"
        gSQL = gSQL & "  FROM BaixaAbastecimento"
        gSQL = gSQL & " WHERE Empresa = " & pEmpresa
        gSQL = gSQL & "   AND [String da Automacao] = " & preparaTexto(pStringAutomacao)
        ExisteStringAutomacao = Localizar(1)
        rsMovAbastecimento.Close
        Set rsMovAbastecimento = Nothing
    End If
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, Hora DESC, Bico DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsMovAbastecimento.Close
    Set rsMovAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Movimento_Abastecimento ( Empresa, Bico, Data, Hora, [Tempo de Abastecimento], "
    gSQL = gSQL & "[Codigo do Produto], [Valor Unitario], Quantidade, [Valor Total], Acerto, "
    gSQL = gSQL & "[Numero do Cupom], [Codigo da ECF], [Numero do Abastecimento], Encerrante, [String da Automacao], "
    gSQL = gSQL & "[Documento Gerado], [Complemento Documento Gerado], [Valor do Desconto], [Encerrante Inicial], [Tipo de Combustivel], "
    gSQL = gSQL & "[Codigo do Funcionario], Periodo, [Valor Desconto Nivel Preco], [Valor Acrescimo Nivel Preco] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mBico, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlHora(1, mHora, ", ")
    Call sqlTexto(1, mTempoAbastecimento, ", ")
    Call sqlNumero(1, mCodigoProduto, ", ")
    Call sqlValor(1, mValorUnitario, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlValor(1, mValorTotal, ", ")
    Call sqlBoolean(1, mAcerto, ", ")
    Call sqlNumero(1, mNumeroCupom, ", ")
    Call sqlNumero(1, mCodigoECF, ", ")
    Call sqlTexto(1, mNumeroAbastecimento, ", ")
    Call sqlValor(1, mEncerrante, ", ")
    Call sqlTexto(1, mStringAutomacao, ", ")
    Call sqlTexto(1, mDocumentoGerado, ", ")
    Call sqlTexto(1, mComplementoDocumentoGerado, ", ")
    Call sqlValor(1, mValorDesconto, ", ")
    Call sqlValor(1, mEncerranteInicial, ", ")
    Call sqlTexto(1, mTipoCombustivel, ", ")
    Call sqlNumero(1, mCodigoFuncionario, ", ")
    Call sqlNumero(1, mPeriodo, ", ")
    
    If IsNull(mValorDescontoNivelPreco) Then
        mValorDescontoNivelPreco = 0
    End If
    Call sqlValor(1, mValorDescontoNivelPreco, ", ")
    
    If IsNull(mValorAcrescimoNivelPreco) Then
        mValorAcrescimoNivelPreco = 0
    End If
    Call sqlValor(1, mValorAcrescimoNivelPreco, " )")
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pHora As Date, ByVal pBico As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Movimento_Abastecimento SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", Bico = ", mBico)
    Call sqlData(2, ", Data = ", mData)
    Call sqlHora(2, ", Hora = ", mHora)
    Call sqlTexto(2, ", [Tempo de Abastecimento] = ", mTempoAbastecimento)
    Call sqlNumero(2, ", [Codigo do Produto] = ", mCodigoProduto)
    Call sqlValor(2, ", [Valor Unitario] = ", mValorUnitario)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlValor(2, ", [Valor Total] = ", mValorTotal)
    Call sqlBoolean(2, ", Acerto = ", mAcerto)
    Call sqlNumero(2, ", [Numero do Cupom] = ", mNumeroCupom)
    Call sqlNumero(2, ", [Codigo da ECF] = ", mCodigoECF)
    Call sqlTexto(2, ", [Numero do Abastecimento] = ", mNumeroAbastecimento)
    Call sqlValor(2, ", Encerrante = ", mEncerrante)
    Call sqlTexto(2, ", [String da Automacao] = ", mStringAutomacao)
    Call sqlTexto(2, ", [Documento Gerado] = ", mDocumentoGerado)
    Call sqlTexto(2, ", [Complemento Documento Gerado] = ", mComplementoDocumentoGerado)
    Call sqlValor(2, ", [Valor do Desconto] = ", mValorDesconto)
    Call sqlValor(2, ", [Encerrante Inicial] = ", mEncerranteInicial)
    Call sqlTexto(2, ", [Tipo de Combustivel] = ", mTipoCombustivel)
    Call sqlNumero(2, ", [Codigo do Funcionario] = ", mCodigoFuncionario)
    Call sqlNumero(2, ", Periodo = ", mPeriodo)
    Call sqlValor(2, ", [Valor Desconto Nivel Preco] = ", mValorDescontoNivelPreco)
    Call sqlValor(2, ", [Valor Acrescimo Nivel Preco] = ", mValorAcrescimoNivelPreco)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND " & preparaHoraConsulta("Hora", "=", pHora)
    gSQL = gSQL & " AND Bico = " & pBico
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function AlterarDesconto(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pHora As Date, ByVal pBico As Integer, ByVal pValorDesconto As Currency, ByVal pNovaQuantidade As Currency) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    AlterarDesconto = False
    gSQL = "UPDATE Movimento_Abastecimento SET "
    Call sqlValor(2, "Quantidade = ", pNovaQuantidade)
    Call sqlValor(2, ", [Valor do Desconto] = ", pValorDesconto)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND " & preparaHoraConsulta("Hora", "=", pHora)
    gSQL = gSQL & " AND Bico = " & pBico
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlterarDesconto = True
    End If
    Exit Function

Err_Alterar:
End Function

Public Function Excluir(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pHora As Date, ByVal pBico As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Movimento_Abastecimento"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND " & preparaHoraConsulta("Hora", "=", pHora)
    gSQL = gSQL & " AND Bico = " & pBico
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


Public Function DescarregarAbastecimento(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pCodigoECF As Integer, ByVal pDocumentoGerado As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_DescarregarAbastecimento

    DescarregarAbastecimento = False
    gSQL = "UPDATE Movimento_Abastecimento SET "
    Call sqlBoolean(2, "Acerto = ", "True")
    Call sqlNumero(2, ", [Numero do Cupom] = ", "1")
    Call sqlNumero(2, ", [Codigo da ECF] = ", pCodigoECF)
    Call sqlTexto(2, ", [Documento Gerado] = ", pDocumentoGerado)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Acerto = " & preparaBooleano(False)
    gSQL = gSQL & " AND [Numero do Cupom] = " & 0
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        DescarregarAbastecimento = True
    End If
    Exit Function

Err_DescarregarAbastecimento:
    MsgBox "Não possível descarregar os abastecimentos pendentes!" & Chr(10) & "Favor entrar em contato com o suporte!", vbCritical, "Erro de Integridade!"
End Function


Public Function DescarregarAbastecimentoFuncionario(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pCodigoECF As Integer, ByVal pDocumentoGerado As String, ByVal pCodigoFuncionario As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_DescarregarAbastecimento

    DescarregarAbastecimentoFuncionario = False
    gSQL = "UPDATE Movimento_Abastecimento SET "
    Call sqlBoolean(2, "Acerto = ", "True")
    Call sqlNumero(2, ", [Numero do Cupom] = ", "1")
    Call sqlNumero(2, ", [Codigo da ECF] = ", pCodigoECF)
    Call sqlTexto(2, ", [Documento Gerado] = ", pDocumentoGerado)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Acerto = " & preparaBooleano(False)
    gSQL = gSQL & " AND [Numero do Cupom] = " & 0
    gSQL = gSQL & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        DescarregarAbastecimentoFuncionario = True
    End If
    Exit Function

Err_DescarregarAbastecimento:
    MsgBox "Não possível descarregar os abastecimentos pendentes!" & Chr(10) & "Favor entrar em contato com o suporte!", vbCritical, "Erro de Integridade!"
End Function
Public Function DefineAbastecimentoVinculadoAfericao(ByVal pEmpresa As Integer, ByVal pDataAbastecimento As Date, ByVal pHoraAbastecimento As Date, ByVal pBico As Integer, ByVal pPeriodo As Integer, ByVal pCodigoFuncionario As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_DefineAbastecimentoVinculadoAfericao

    DefineAbastecimentoVinculadoAfericao = False
    gSQL = "UPDATE Movimento_Abastecimento SET "
    Call sqlTexto(2, " [Complemento Documento Gerado] = ", "AFERICAO")
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pDataAbastecimento)
    gSQL = gSQL & " AND Hora = " & preparaHora(pHoraAbastecimento)
    If pCodigoFuncionario > 0 Then
        gSQL = gSQL & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    End If
    gSQL = gSQL & " AND Bico = " & pBico
    gSQL = gSQL & " AND [Complemento Documento Gerado] = " & preparaTexto("AFTEMP")
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        DefineAbastecimentoVinculadoAfericao = True
    End If
    Exit Function

Err_DefineAbastecimentoVinculadoAfericao:
    MsgBox "Não possível alterar abastecimento vinculado à aferição!" & Chr(10) & "Favor entrar em contato com o suporte!", vbCritical, "Erro de Integridade!"
End Function
Public Function DesvincularAbastecimentoComAfericao(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pHora As Date, ByVal pCodigoFuncionario As Integer, pBico As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_DesvincularAbastecimentoComAfericao

    DesvincularAbastecimentoComAfericao = False
    gSQL = "UPDATE Movimento_Abastecimento SET "
    Call sqlTexto(2, " [Complemento Documento Gerado] = ", "")
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Hora = " & preparaHora(pHora)
    gSQL = gSQL & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    gSQL = gSQL & " AND Bico = " & pBico
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        DesvincularAbastecimentoComAfericao = True
    End If
    Exit Function

Err_DesvincularAbastecimentoComAfericao:
    MsgBox "Não possível desvincular o abastecimento à aferição!" & Chr(10) & "Favor entrar em contato com o suporte!", vbCritical, "Erro de Integridade!"
End Function
Public Function LiberaAbastecimentoAfericaoTemporaria(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pCodigoFuncionario As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_LiberaAbastecimentoAfericaoTemporaria

    LiberaAbastecimentoAfericaoTemporaria = False
    gSQL = "UPDATE Movimento_Abastecimento SET "
    Call sqlBoolean(2, "Acerto = ", "False")
    Call sqlNumero(2, ", [Numero do Cupom] = ", "0")
    Call sqlNumero(2, ", [Codigo da ECF] = ", 0)
    Call sqlTexto(2, ", [Documento Gerado] = ", "")
    Call sqlTexto(2, " [Complemento Documento Gerado] = ", "")
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Complemento Documento Gerado] = " & preparaTexto("AFTEMP")
    
    If pCodigoFuncionario > 0 Then
        gSQL = gSQL & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    End If
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        LiberaAbastecimentoAfericaoTemporaria = True
    End If
    Exit Function

Err_LiberaAbastecimentoAfericaoTemporaria:
    MsgBox "Não possível liberar o abastecimento com aferição temporária!" & Chr(10) & "Favor entrar em contato com o suporte!", vbCritical, "Erro de Integridade!"
End Function
Public Function LiberaAbastecimentoAfericao(ByVal pEmpresa As Integer, ByVal pDataAbastecimento As Date, ByVal pHoraAbastecimento As Date, ByVal pBico As Integer, ByVal pCodigoFuncionario As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_LiberaAbastecimentoAfericao

    LiberaAbastecimentoAfericao = False
    gSQL = "UPDATE Movimento_Abastecimento SET "
    Call sqlBoolean(2, "Acerto = ", "False")
    Call sqlNumero(2, ", [Numero do Cupom] = ", "0")
    Call sqlNumero(2, ", [Codigo da ECF] = ", 0)
    Call sqlTexto(2, ", [Documento Gerado] = ", "")
    Call sqlTexto(2, ", [Complemento Documento Gerado] = ", "")
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pDataAbastecimento)
    gSQL = gSQL & " AND Hora = " & preparaHora(pHoraAbastecimento)
    gSQL = gSQL & " AND Bico = " & pBico
    gSQL = gSQL & " AND ([Complemento Documento Gerado] = " & preparaTexto("AFERICAO") & " OR [Complemento Documento Gerado] = " & preparaTexto("AFTEMP") & ")"
    
    
    If pCodigoFuncionario > 0 Then
        gSQL = gSQL & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
    End If
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        LiberaAbastecimentoAfericao = True
    End If
    Exit Function

Err_LiberaAbastecimentoAfericao:
    MsgBox "Não possível liberar o abastecimento da aferição!" & Chr(10) & "Favor entrar em contato com o suporte!", vbCritical, "Erro de Integridade!"
End Function

'Rotina para voltar um abastecimento impresso no ecf
'Por Motivo do ecf ter sido cancelado
Public Function VoltaEcfCancelado(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumeroCupom As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Interno

    VoltaEcfCancelado = False
    gSQL = "UPDATE Movimento_Abastecimento SET "
    Call sqlBoolean(2, "Acerto = ", "False")
    Call sqlNumero(2, ", [Numero do Cupom] = ", 0)
    Call sqlNumero(2, ", [Codigo da ECF] = ", 0)
    Call sqlTexto(2, ", [Documento Gerado] = ", "")
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Numero do Cupom] = " & pNumeroCupom
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        VoltaEcfCancelado = True
    End If
    Exit Function

Err_Interno:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal pCondicao As String, ByVal pTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Bico, Data, Hora, [Tempo de Abastecimento], [Codigo do Produto], [Valor Unitario], Quantidade, [Valor Total], Acerto, [Numero do Cupom], [Codigo da ECF], [Numero do Abastecimento], Encerrante, [String da Automacao], [Documento Gerado], [Complemento Documento Gerado], [Valor do Desconto], [Encerrante Inicial], [Tipo de Combustivel], [Codigo do Funcionario], Periodo, [Valor Desconto Nivel Preco], [Valor Acrescimo Nivel Preco] FROM Movimento_Abastecimento" & " " & pCondicao & " " & pTipoOrdem)
End Function





Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovAbastecimento = New adodb.Recordset
    rsMovAbastecimento.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovAbastecimento.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovAbastecimento.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovAbastecimento("Empresa").Value
    Bico = rsMovAbastecimento("Bico").Value
    Data = rsMovAbastecimento("Data").Value
    Hora = rsMovAbastecimento("Hora").Value
    TempoAbastecimento = rsMovAbastecimento("Tempo de Abastecimento").Value
    CodigoProduto = rsMovAbastecimento("Codigo do Produto").Value
    ValorUnitario = rsMovAbastecimento("Valor Unitario").Value
    Quantidade = rsMovAbastecimento("Quantidade").Value
    ValorTotal = rsMovAbastecimento("Valor Total").Value
    Acerto = rsMovAbastecimento("Acerto").Value
    NumeroCupom = rsMovAbastecimento("Numero do Cupom").Value
    CodigoECF = rsMovAbastecimento("Codigo da ECF").Value
    NumeroAbastecimento = rsMovAbastecimento("Numero do Abastecimento").Value
    Encerrante = rsMovAbastecimento("Encerrante").Value
    StringAutomacao = rsMovAbastecimento("String da Automacao").Value
    DocumentoGerado = rsMovAbastecimento("Documento Gerado").Value
    ComplementoDocumentoGerado = rsMovAbastecimento("Complemento Documento Gerado").Value
    ValorDesconto = rsMovAbastecimento("Valor do Desconto").Value
    EncerranteInicial = rsMovAbastecimento("Encerrante Inicial").Value
    TipoCombustivel = rsMovAbastecimento("Tipo de Combustivel").Value
    CodigoFuncionario = rsMovAbastecimento("Codigo do Funcionario").Value
    Periodo = rsMovAbastecimento("Periodo").Value
    If IsNull(rsMovAbastecimento("Valor Desconto Nivel Preco").Value) Then
        ValorDescontoNivelPreco = 0
    Else
        ValorDescontoNivelPreco = rsMovAbastecimento("Valor Desconto Nivel Preco").Value
    End If
    
    If IsNull(rsMovAbastecimento("Valor Acrescimo Nivel Preco").Value) Then
        ValorAcrescimoNivelPreco = 0
    Else
        ValorAcrescimoNivelPreco = rsMovAbastecimento("Valor Acrescimo Nivel Preco").Value
    End If
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

