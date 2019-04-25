VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBaixaAbastecimento"
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
Private mNumeroAbastecimento As String
Private mEncerrante As String
Private mStringAutomacao As String
Private mCodigoECF As Integer
Private mNumeroCupom As Long
Private mDataBaixa As Date
Private mHoraBaixa As Date
Private mDocumentoGerado As String
Private mComplementoDocumentoGerado As String
Private mValorDesconto As Currency
Private mAcerto As Boolean
Private mEncerranteInicial As Currency
Private mTipoCombustivel As String
Private mCodigoFuncionario As Integer
Private mPeriodo As Integer
Private mValorDescontoNivelPreco As Currency
Private mValorAcrescimoNivelPreco As Currency


Private rsBaixaAbastecimento As New adodb.Recordset


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


Public Property Get NumeroAbastecimento() As String
    NumeroAbastecimento = mNumeroAbastecimento
End Property
Public Property Let NumeroAbastecimento(ByVal Valor As String)
    mNumeroAbastecimento = Valor
End Property


Public Property Get Encerrante() As String
    Encerrante = mEncerrante
End Property
Public Property Let Encerrante(ByVal Valor As String)
    mEncerrante = Valor
End Property


Public Property Get StringAutomacao() As String
    StringAutomacao = mStringAutomacao
End Property
Public Property Let StringAutomacao(ByVal Valor As String)
    mStringAutomacao = Valor
End Property


Public Property Get CodigoECF() As Integer
    CodigoECF = mCodigoECF
End Property
Public Property Let CodigoECF(ByVal Valor As Integer)
    mCodigoECF = Valor
End Property


Public Property Get NumeroCupom() As Long
    NumeroCupom = mNumeroCupom
End Property
Public Property Let NumeroCupom(ByVal Valor As Long)
    mNumeroCupom = Valor
End Property


Public Property Get DataBaixa() As Date
    DataBaixa = mDataBaixa
End Property
Public Property Let DataBaixa(ByVal Valor As Date)
    mDataBaixa = Valor
End Property


Public Property Get HoraBaixa() As Date
    HoraBaixa = mHoraBaixa
End Property
Public Property Let HoraBaixa(ByVal Valor As Date)
    mHoraBaixa = Valor
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


Public Property Get Acerto() As Boolean
    Acerto = mAcerto
End Property
Public Property Let Acerto(ByVal Valor As Boolean)
    mAcerto = Valor
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
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pBico As Integer, ByVal pData As Date, ByVal pHora As Date) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Bico = " & pBico
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND Hora = " & preparaHora(pHora)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsBaixaAbastecimento.Close
    Set rsBaixaAbastecimento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO BaixaAbastecimento ( Empresa, Bico, Data, Hora, [Tempo de Abastecimento], "
    gSQL = gSQL & "[Codigo do Produto], [Valor Unitario], Quantidade, [Valor Total], [Numero do Abastecimento], "
    gSQL = gSQL & "Encerrante, [String da Automacao], [Codigo da ECF], [Numero do Cupom], [Data da Baixa], "
    gSQL = gSQL & "[Hora da Baixa], [Documento Gerado], [Complemento Documento Gerado], [Valor do Desconto], "
    gSQL = gSQL & "Acerto, [Encerrante Inicial], [Tipo de Combustivel], [Codigo do Funcionario], "
    gSQL = gSQL & "Periodo, [Valor Desconto Nivel Preco], [Valor Acrescimo Nivel Preco] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mBico, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlHora(1, mHora, ", ")
    Call sqlTexto(1, mTempoAbastecimento, ", ")
    Call sqlNumero(1, mCodigoProduto, ", ")
    Call sqlValor(1, mValorUnitario, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlValor(1, mValorTotal, ", ")
    Call sqlTexto(1, mNumeroAbastecimento, ", ")
    Call sqlTexto(1, mEncerrante, ", ")
    Call sqlTexto(1, mStringAutomacao, ", ")
    Call sqlNumero(1, mCodigoECF, ", ")
    Call sqlNumero(1, mNumeroCupom, ", ")
    Call sqlData(1, mDataBaixa, ", ")
    Call sqlHora(1, mHoraBaixa, ", ")
    Call sqlTexto(1, mDocumentoGerado, ", ")
    Call sqlTexto(1, mComplementoDocumentoGerado, ", ")
    Call sqlValor(1, mValorDesconto, ", ")
    Call sqlBoolean(1, mAcerto, ", ")
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
    Else
        Call CriaLogCadastroDll2("Classe:cBaixaAbastecimento - Método:Incluir", "Não foi possível incluir registro.", gSQL)
    End If
    Exit Function

Err_Incluir:
    Call CriaLogCadastroDll2("Classe:cBaixaAbastecimento - Método:Incluir", Err.Description, gSQL)
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pBico As Integer, ByVal pData As Date, ByVal pHora As Date) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE BaixaAbastecimento SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", Bico = ", mBico)
    Call sqlData(2, ", Data = ", mData)
    Call sqlHora(2, ", Hora = ", mHora)
    Call sqlTexto(2, ", [Tempo de Abastecimento] = ", mTempoAbastecimento)
    Call sqlNumero(2, ", [Codigo do Produto] = ", mCodigoProduto)
    Call sqlValor(2, ", [Valor Unitario] = ", mValorUnitario)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlValor(2, ", [Valor Total] = ", mValorTotal)
    Call sqlTexto(2, ", [Numero do Abastecimento] = ", mNumeroAbastecimento)
    Call sqlTexto(2, ", Encerrante = ", mEncerrante)
    Call sqlTexto(2, ", [String da Automacao] = ", mStringAutomacao)
    Call sqlNumero(2, ", [Codigo da ECF] = ", mCodigoECF)
    Call sqlNumero(2, ", [Numero do Cupom] = ", mNumeroCupom)
    Call sqlData(2, ", [Data da Baixa] = ", mDataBaixa)
    Call sqlHora(2, ", [Hora da Baixa] = ", mHoraBaixa)
    Call sqlTexto(2, ", [Documento Gerado] = ", mDocumentoGerado)
    Call sqlTexto(2, ", [Complemento Documento Gerado] = ", mComplementoDocumentoGerado)
    Call sqlValor(2, ", [Valor do Desconto] = ", mValorDesconto)
    Call sqlBoolean(2, ", Acerto = ", mAcerto)
    Call sqlValor(2, ", [Encerrante Inicial] = ", mEncerranteInicial)
    Call sqlTexto(2, ", [Tipo de Combustivel] = ", mTipoCombustivel)
    Call sqlNumero(2, ", [Codigo do Funcionario] = ", mCodigoFuncionario)
    Call sqlNumero(2, ", Periodo = ", mPeriodo)
    Call sqlValor(2, ", [Valor Desconto Nivel Preco] = ", mValorDescontoNivelPreco)
    Call sqlValor(2, ", [Valor Acrescimo Nivel Preco] = ", mValorAcrescimoNivelPreco)
    
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Bico = " & pBico
    gSQL = gSQL & "   AND Data = " & preparaData(pData)
    gSQL = gSQL & "   AND Hora = " & preparaHora(pHora)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pBico As Integer, ByVal pData As Date, ByVal pHora As Date) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM BaixaAbastecimento"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Bico = " & pBico
    gSQL = gSQL & "   AND Data = " & preparaData(pData)
    gSQL = gSQL & "   AND Hora = " & preparaHora(pHora)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Bico, Data, Hora, [Tempo de Abastecimento], [Codigo do Produto], [Valor Unitario], Quantidade, [Valor Total], [Numero do Abastecimento], Encerrante, [String da Automacao], [Codigo da ECF], [Numero do Cupom], [Data da Baixa], [Hora da Baixa], [Documento Gerado], [Complemento Documento Gerado], [Valor do Desconto], Acerto, [Encerrante Inicial], [Tipo de Combustivel], [Codigo do Funcionario], Periodo, [Valor Desconto Nivel Preco], [Valor Acrescimo Nivel Preco] FROM BaixaAbastecimento" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsBaixaAbastecimento = New adodb.Recordset
    rsBaixaAbastecimento.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsBaixaAbastecimento.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsBaixaAbastecimento.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsBaixaAbastecimento("Empresa").Value
    Bico = rsBaixaAbastecimento("Bico").Value
    Data = rsBaixaAbastecimento("Data").Value
    Hora = rsBaixaAbastecimento("Hora").Value
    TempoAbastecimento = rsBaixaAbastecimento("Tempo de Abastecimento").Value
    CodigoProduto = rsBaixaAbastecimento("Codigo do Produto").Value
    ValorUnitario = rsBaixaAbastecimento("Valor Unitario").Value
    Quantidade = rsBaixaAbastecimento("Quantidade").Value
    ValorTotal = rsBaixaAbastecimento("Valor Total").Value
    NumeroAbastecimento = rsBaixaAbastecimento("Numero do Abastecimento").Value
    Encerrante = rsBaixaAbastecimento("Encerrante").Value
    StringAutomacao = rsBaixaAbastecimento("String da Automacao").Value
    CodigoECF = rsBaixaAbastecimento("Codigo da ECF").Value
    NumeroCupom = rsBaixaAbastecimento("Numero do Cupom").Value
    If IsNull(rsBaixaAbastecimento("Data da Baixa").Value) Then
        DataBaixa = "00:00:00"
    Else
        DataBaixa = rsBaixaAbastecimento("Data da Baixa").Value
    End If
    If IsNull(rsBaixaAbastecimento("Hora da Baixa").Value) Then
        HoraBaixa = "00:00:00"
    Else
        HoraBaixa = rsBaixaAbastecimento("Hora da Baixa").Value
    End If
    DocumentoGerado = rsBaixaAbastecimento("Documento Gerado").Value
    ComplementoDocumentoGerado = rsBaixaAbastecimento("Complemento Documento Gerado").Value
    ValorDesconto = rsBaixaAbastecimento("Valor do Desconto").Value
    Acerto = rsBaixaAbastecimento("Acerto").Value
    EncerranteInicial = rsBaixaAbastecimento("Encerrante Inicial").Value
    TipoCombustivel = rsBaixaAbastecimento("Tipo de Combustivel").Value
    CodigoFuncionario = rsBaixaAbastecimento("Codigo do Funcionario").Value
    Periodo = rsBaixaAbastecimento("Periodo").Value
    
    If IsNull(rsBaixaAbastecimento("Valor Desconto Nivel Preco").Value) Then
        ValorDescontoNivelPreco = 0
    Else
        ValorDescontoNivelPreco = rsBaixaAbastecimento("Valor Desconto Nivel Preco").Value
    End If
    
    If IsNull(rsBaixaAbastecimento("Valor Acrescimo Nivel Preco").Value) Then
        ValorAcrescimoNivelPreco = 0
    Else
        ValorAcrescimoNivelPreco = rsBaixaAbastecimento("Valor Acrescimo Nivel Preco").Value
    End If

End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

