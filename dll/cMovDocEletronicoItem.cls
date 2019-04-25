VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovDocEletronicoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mIdEstabelecimento As Integer
Private mDataEmissao As Date
Private mEntrada As Boolean
Private mSaida As Boolean
Private mModelo As String
Private mSerie As String
Private mNumero As String
Private mOrdem As Integer
Private mIdClienteFornecedor As Long
Private mMovimentacaoFisica As Boolean
Private mCFOP As String
Private mIdProduto As Long
Private mValorUnitario As Currency
Private mQuantidade As Currency
Private mValorDesconto As Currency
Private mValorTotalLiquido As Currency
Private mCSTICMS As String
Private mValorBCICMS As Currency
Private mAliquotaICMS As Currency
Private mValorICMS As Currency
Private mCstIcmsSt As String
Private mValorBCICMSST As Currency
Private mAliquotaIcmsSt As Currency
Private mValorICMSST As Currency
Private mCSTIPI As String
Private mValorBcIpi As Currency
Private mAliquotaIpi As Currency
Private mValorIPI As Currency
Private mApuracaoIpiMensal As Boolean
Private mCodigoEnquadramentoIpi As String
Private mCstPis As String
Private mValorBcPis As Currency
Private mQuantidadeBcPis As Currency
Private mAliquotaPis As Currency
Private mValorPis As Currency
Private mCstCofins As String
Private mValorBcCofins As Currency
Private mQuantidadeBcCofins As Currency
Private mAliquotaCofins As Currency
Private mValorCofins As Currency
Private mNumeroTanque As Integer
Private mCancelado As Boolean
Private mDataEntradaSaida As Date
Private mEncerranteFinal As Currency
Private mNumeroBomba As Integer
Private mNumeroBico As Integer
Private mTipoCombustivel As String
Private mEtapaConcluida As Integer
Private mProgramaOrigem As String
Private mPeriodo As String


Private rsMovDocumentoEletronicoItem As New adodb.Recordset
'Private IntegracaoNuvem As New cIntegracaoNuvem


'Inicio das Propriedades da Classe
Public Property Get IdEstabelecimento() As Integer
    IdEstabelecimento = mIdEstabelecimento
End Property
Public Property Let IdEstabelecimento(ByVal Valor As Integer)
    mIdEstabelecimento = Valor
End Property


Public Property Get DataEmissao() As Date
    DataEmissao = mDataEmissao
End Property
Public Property Let DataEmissao(ByVal Valor As Date)
    mDataEmissao = Valor
End Property


Public Property Get Entrada() As Boolean
    Entrada = mEntrada
End Property
Public Property Let Entrada(ByVal Valor As Boolean)
    mEntrada = Valor
End Property
Public Property Get Saida() As Boolean
    Saida = mSaida
End Property
Public Property Let Saida(ByVal Valor As Boolean)
    mSaida = Valor
End Property

Public Property Get Modelo() As String
    Modelo = mModelo
End Property
Public Property Let Modelo(ByVal Valor As String)
    mModelo = Valor
End Property

Public Property Get Serie() As String
    Serie = mSerie
End Property
Public Property Let Serie(ByVal Valor As String)
    mSerie = Valor
End Property

Public Property Get Numero() As String
    mNumero = mNumero
End Property
Public Property Let Numero(ByVal Valor As String)
    mNumero = Valor
End Property


Public Property Get Ordem() As String
    Ordem = mOrdem
End Property
Public Property Let Ordem(ByVal Valor As String)
    mOrdem = Valor
End Property

Public Property Get IdClienteFornecedor() As Long
    IdClienteFornecedor = mIdClienteFornecedor
End Property
Public Property Let IdClienteFornecedor(ByVal Valor As Long)
    mIdClienteFornecedor = Valor
End Property

Public Property Get MovimentacaoFisica() As Boolean
    MovimentacaoFisica = mMovimentacaoFisica
End Property
Public Property Let MovimentacaoFisica(ByVal Valor As Boolean)
    mMovimentacaoFisica = Valor
End Property

Public Property Get CFOP() As String
    CFOP = mCFOP
End Property
Public Property Let CFOP(ByVal Valor As String)
    mCFOP = Valor
End Property

Public Property Get IdProduto() As Long
    IdProduto = mIdProduto
End Property
Public Property Let IdProduto(ByVal Valor As Long)
    mIdProduto = Valor
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


Public Property Get ValorDesconto() As Currency
    ValorDesconto = mValorDesconto
End Property
Public Property Let ValorDesconto(ByVal Valor As Currency)
    mValorDesconto = Valor
End Property

Public Property Get ValorTotalLiquido() As Currency
    ValorTotalLiquido = mValorTotalLiquido
End Property
Public Property Let ValorTotalLiquido(ByVal Valor As Currency)
    mValorTotalLiquido = Valor
End Property

Public Property Get CSTICMS() As String
    CSTICMS = mCSTICMS
End Property
Public Property Let CSTICMS(ByVal Valor As String)
    mCSTICMS = Valor
End Property

Public Property Get ValorBCICMS() As Currency
    ValorBCICMS = mValorBCICMS
End Property
Public Property Let ValorBCICMS(ByVal Valor As Currency)
    mValorBCICMS = Valor
End Property

Public Property Get AliquotaICMS() As Currency
    AliquotaICMS = mAliquotaICMS
End Property
Public Property Let AliquotaICMS(ByVal Valor As Currency)
    mAliquotaICMS = Valor
End Property

Public Property Get ValorICMS() As Currency
    ValorICMS = mValorICMS
End Property
Public Property Let ValorICMS(ByVal Valor As Currency)
    mValorICMS = Valor
End Property

Public Property Get CstIcmsSt() As String
    CstIcmsSt = mCstIcmsSt
End Property
Public Property Let CstIcmsSt(ByVal Valor As String)
    mCstIcmsSt = Valor
End Property

Public Property Get ValorBCICMSST() As Currency
    ValorBCICMSST = mValorBCICMSST
End Property
Public Property Let ValorBCICMSST(ByVal Valor As Currency)
    mValorBCICMSST = Valor
End Property
Public Property Get AliquotaIcmsSt() As Currency
    AliquotaIcmsSt = mAliquotaIcmsSt
End Property
Public Property Let AliquotaIcmsSt(ByVal Valor As Currency)
    mAliquotaIcmsSt = Valor
End Property
Public Property Get ValorICMSST() As Currency
    ValorICMSST = mValorICMSST
End Property
Public Property Let ValorICMSST(ByVal Valor As Currency)
    mValorICMSST = Valor
End Property
Public Property Get CSTIPI() As String
    CSTIPI = mCSTIPI
End Property
Public Property Let CSTIPI(ByVal Valor As String)
    mCSTIPI = Valor
End Property
Public Property Get ValorBcIpi() As Currency
    ValorBcIpi = mValorBcIpi
End Property
Public Property Let ValorBcIpi(ByVal Valor As Currency)
    mValorBcIpi = Valor
End Property
Public Property Get AliquotaIpi() As Currency
    AliquotaIpi = mAliquotaIpi
End Property
Public Property Let AliquotaIpi(ByVal Valor As Currency)
    mAliquotaIpi = Valor
End Property
Public Property Get ValorIPI() As Currency
    ValorIPI = mValorIPI
End Property
Public Property Let ValorIPI(ByVal Valor As Currency)
    mValorIPI = Valor
End Property
Public Property Get ApuracaoIpiMensal() As Boolean
    ApuracaoIpiMensal = mApuracaoIpiMensal
End Property
Public Property Let ApuracaoIpiMensal(ByVal Valor As Boolean)
    mApuracaoIpiMensal = Valor
End Property
Public Property Get CodigoEnquadramentoIpi() As String
    CodigoEnquadramentoIpi = mCodigoEnquadramentoIpi
End Property
Public Property Let CodigoEnquadramentoIpi(ByVal Valor As String)
    mCodigoEnquadramentoIpi = Valor
End Property
Public Property Get CstPis() As String
    CstPis = mCstPis
End Property
Public Property Let CstPis(ByVal Valor As String)
    mCstPis = Valor
End Property
Public Property Get ValorBcPis() As Currency
    ValorBcPis = mValorBcPis
End Property
Public Property Let ValorBcPis(ByVal Valor As Currency)
    mValorBcPis = Valor
End Property
Public Property Get QuantidadeBcPis() As Currency
    QuantidadeBcPis = mQuantidadeBcPis
End Property
Public Property Let QuantidadeBcPis(ByVal Valor As Currency)
    mQuantidadeBcPis = Valor
End Property
Public Property Get AliquotaPis() As Currency
    AliquotaPis = mAliquotaPis
End Property
Public Property Let AliquotaPis(ByVal Valor As Currency)
    mAliquotaPis = Valor
End Property
Public Property Get ValorPis() As Currency
    ValorPis = mValorPis
End Property
Public Property Let ValorPis(ByVal Valor As Currency)
    mValorPis = Valor
End Property
Public Property Get CstCofins() As String
    CstCofins = mCstCofins
End Property
Public Property Let CstCofins(ByVal Valor As String)
    mCstCofins = Valor
End Property
Public Property Get ValorBcCofins() As Currency
    ValorBcCofins = mValorBcCofins
End Property
Public Property Let ValorBcCofins(ByVal Valor As Currency)
    mValorBcCofins = Valor
End Property
Public Property Get QuantidadeBcCofins() As Currency
    QuantidadeBcCofins = mQuantidadeBcCofins
End Property
Public Property Let QuantidadeBcCofins(ByVal Valor As Currency)
    mQuantidadeBcCofins = Valor
End Property
Public Property Get AliquotaCofins() As Currency
    AliquotaCofins = mAliquotaCofins
End Property
Public Property Let AliquotaCofins(ByVal Valor As Currency)
    mAliquotaCofins = Valor
End Property
Public Property Get ValorCofins() As Currency
    ValorCofins = mValorCofins
End Property
Public Property Let ValorCofins(ByVal Valor As Currency)
    mValorCofins = Valor
End Property
Public Property Get NumeroTanque() As Integer
    NumeroTanque = mNumeroTanque
End Property
Public Property Let NumeroTanque(ByVal Valor As Integer)
    mNumeroTanque = Valor
End Property
Public Property Get Cancelado() As Integer
    Cancelado = mCancelado
End Property
Public Property Let Cancelado(ByVal Valor As Integer)
    mCancelado = Valor
End Property
Public Property Get DataEntradaSaida() As Date
    DataEntradaSaida = mDataEntradaSaida
End Property
Public Property Let DataEntradaSaida(ByVal Valor As Date)
    mDataEntradaSaida = Valor
End Property
Public Property Get EncerranteFinal() As Currency
    EncerranteFinal = mEncerranteFinal
End Property
Public Property Let EncerranteFinal(ByVal Valor As Currency)
    mEncerranteFinal = Valor
End Property
Public Property Get NumeroBomba() As Integer
    NumeroBomba = mNumeroBomba
End Property
Public Property Let NumeroBomba(ByVal Valor As Integer)
    mNumeroBomba = Valor
End Property
Public Property Get NumeroBico() As Integer
    NumeroBico = mNumeroBico
End Property
Public Property Let NumeroBico(ByVal Valor As Integer)
    mNumeroBico = Valor
End Property
Public Property Get TipoCombustivel() As String
    TipoCombustivel = mTipoCombustivel
End Property
Public Property Let TipoCombustivel(ByVal Valor As String)
    mTipoCombustivel = Valor
End Property
Public Property Get EtapaConcluida() As Integer
    EtapaConcluida = mEtapaConcluida
End Property
Public Property Let EtapaConcluida(ByVal Valor As Integer)
    mEtapaConcluida = Valor
End Property
Public Property Get ProgramaOrigem() As String
    ProgramaOrigem = mProgramaOrigem
End Property
Public Property Let ProgramaOrigem(ByVal Valor As String)
    mProgramaOrigem = Valor
End Property
Public Property Get Periodo() As String
    Periodo = mPeriodo
End Property
Public Property Let Periodo(ByVal Valor As String)
    mPeriodo = Valor
End Property



Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False

    gSQL = "INSERT INTO MovimentoDocumentoEletronicoItem (IdEstabelecimento_MovDEItem,DataEmissao_MovDEItem,Entrada_MovDEItem,Saida_MovDEItem,Modelo_MovDEItem,Serie_MovDEItem,Numero_MovDEItem"
    gSQL = gSQL & ",Ordem_MovDEItem,IdClienteFornecedor_MovDEItem,MovimentacaoFisica_MovDEItem,CFOP_MovDEItem,IdProduto_MovDEItem,ValorUnitario_MovDEItem"
    gSQL = gSQL & ",Quantidade_MovDEItem,ValorDesconto_MovDEItem,ValorTotalLiquido_MovDEItem,CstIcms_MovDEItem,ValorBcIcms_MovDEItem,AliquotaIcms_MovDEItem"
    gSQL = gSQL & ",ValorIcms_MovDEItem,CstIcmsSt_MovDEItem,ValorBcIcmsSt_MovDEItem,AliquotaIcmsSt_MovDEItem,ValorIcmsSt_MovDEItem,CstIpi_MovDEItem"
    gSQL = gSQL & ",ValorBcIpi_MovDEItem,AliquotaIpi_MovDEItem,ValorIpi_MovDEItem,ApuracaoIpiMensal_MovDEItem,CodigoEnquadramentoIpi_MovDEItem"
    gSQL = gSQL & ",CstPis_MovDEItem,ValorBcPis_MovDEItem,QuantidadeBcPis_MovDEItem,AliquotaPis_MovDEItem,ValorPis_MovDEItem,CstCofins_MovDEItem"
    gSQL = gSQL & ",ValorBcCofins_MovDEItem,QuantidadeBcCofins_MovDEItem,AliquotaCofins_MovDEItem,ValorCofins_MovDEItem,NumeroTanque_MovDEItem"
    gSQL = gSQL & ",Cancelado_MovDEItem,DataEntradaSaida_MovDEItem, EncerranteFinal_MovDEItem, NumeroBomba_MovDEItem, NumeroBico_MovDEItem"
    gSQL = gSQL & ", TipoCombustivel_MovDEItem, EtapaConcluida_MovDEItem, ProgramaOrigem_MovDEItem, Periodo_MovDEItem ) VALUES ("
    Call sqlNumero(1, mIdEstabelecimento, ", ")
    Call sqlData(1, mDataEmissao, ", ")
    Call sqlBoolean(1, mEntrada, ", ")
    Call sqlBoolean(1, mSaida, ", ")
    Call sqlTexto(1, mModelo, ", ")
    Call sqlTexto(1, mSerie, ", ")
    Call sqlTexto(1, mNumero, ", ")
    Call sqlNumero(1, mOrdem, ", ")
    Call sqlNumero(1, mIdClienteFornecedor, ", ")
    Call sqlBoolean(1, mMovimentacaoFisica, ", ")
    Call sqlTexto(1, mCFOP, ", ")
    Call sqlNumero(1, mIdProduto, ", ")
    Call sqlValor(1, mValorUnitario, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlValor(1, mValorDesconto, ", ")
    Call sqlValor(1, mValorTotalLiquido, ", ")
    Call sqlTexto(1, mCSTICMS, ", ")
    Call sqlValor(1, mValorBCICMS, ", ")
    Call sqlValor(1, mAliquotaICMS, ", ")
    Call sqlValor(1, mValorICMS, ", ")
    Call sqlTexto(1, mCstIcmsSt, ", ")
    Call sqlValor(1, mValorBCICMSST, ", ")
    Call sqlValor(1, mAliquotaIcmsSt, ", ")
    Call sqlValor(1, mValorICMSST, ", ")
    Call sqlTexto(1, mCSTIPI, ", ")
    Call sqlValor(1, mValorBcIpi, ", ")
    Call sqlValor(1, mAliquotaIpi, ", ")
    Call sqlValor(1, mValorIPI, ", ")
    Call sqlBoolean(1, mApuracaoIpiMensal, ", ")
    Call sqlTexto(1, mCodigoEnquadramentoIpi, ", ")
    Call sqlTexto(1, mCstPis, ", ")
    Call sqlValor(1, mValorBcPis, ", ")
    Call sqlValor(1, mQuantidadeBcPis, ", ")
    Call sqlValor(1, mAliquotaPis, ", ")
    Call sqlValor(1, mValorPis, ", ")
    Call sqlTexto(1, mCstCofins, ", ")
    Call sqlValor(1, mValorBcCofins, ", ")
    Call sqlValor(1, mQuantidadeBcCofins, ", ")
    Call sqlValor(1, mAliquotaCofins, ", ")
    Call sqlValor(1, mValorCofins, ", ")
    Call sqlNumero(1, mNumeroTanque, ", ")
    Call sqlBoolean(1, mCancelado, ", ")
    Call sqlData(1, mDataEntradaSaida, ", ")
    Call sqlValor(1, mEncerranteFinal, ", ")
    Call sqlNumero(1, mNumeroBomba, ", ")
    Call sqlNumero(1, mNumeroBico, ", ")
    Call sqlTexto(1, mTipoCombustivel, ", ")
    Call sqlNumero(1, mEtapaConcluida, ", ")
    Call sqlTexto(1, mProgramaOrigem, ", ")
    Call sqlTexto(1, mPeriodo, " )")
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    
    If lRecordsAffected > 0 Then
       Incluir = True
    '   Call IncluiIntegracaoNuvem("INCLUIR", mEmpresa, mCodigoECF, mNumeroCupom, mData, mOrdem)
    End If
    Exit Function
    
Err_Incluir:
End Function

Public Function Alterar(ByVal pIdEstabelecimento As Integer, ByVal pDataEmissao As Date, ByVal pEntrada As Boolean, ByVal pSaida As Boolean, ByVal pModelo As String, ByVal pSerie As String, ByVal pNumero As String, pOrdem As Integer) As Boolean
    Dim lRecordsAffected As Long
    
    On Error GoTo Err_Alterar
    
    Alterar = False
    
    gSQL = "UPDATE MovimentoDocumentoEletronicoItem SET "
    
    Call sqlNumero(2, "IdEstabelecimento_MovDEItem = ", mIdEstabelecimento)
    Call sqlData(2, "DataEmissao_MovDEItem = ", mDataEmissao)
    Call sqlBoolean(2, "Entrada_MovDEItem = ", mEntrada)
    Call sqlBoolean(2, "Saida_MovDEItem = ", mSaida)
    Call sqlTexto(2, "Modelo_MovDEItem = ", mModelo)
    Call sqlTexto(2, "Serie_MovDEItem = ", mSerie)
    Call sqlTexto(2, "Numero_MovDEItem = ", mNumero)
    Call sqlNumero(2, "Ordem_MovDEItem = ", mOrdem)
    Call sqlNumero(2, "IdClienteFornecedor_MovDEItem = ", mIdClienteFornecedor)
    Call sqlBoolean(2, "MovimentacaoFisica_MovDEItem = ", mMovimentacaoFisica)
    Call sqlTexto(2, "CFOP_MovDEItem = ", mCFOP)
    Call sqlNumero(2, "IdProduto_MovDEItem = ", mIdProduto)
    Call sqlValor(2, "ValorUnitario_MovDEItem = ", mValorUnitario)
    Call sqlValor(2, "Quantidade_MovDEItem = ", mQuantidade)
    Call sqlValor(2, "ValorDesconto_MovDEItem = ", mValorDesconto)
    Call sqlValor(2, "ValorTotalLiquido_MovDEItem = ", mValorTotalLiquido)
    Call sqlTexto(2, "CSTICMS_MovDEItem = ", mCSTICMS)
    Call sqlValor(2, "ValorBCICMS_MovDEItem = ", mValorBCICMS)
    Call sqlValor(2, "AliquotaICMS_MovDEItem = ", mAliquotaICMS)
    Call sqlValor(2, "ValorICMS_MovDEItem = ", mValorICMS)
    Call sqlTexto(2, "CstIcmsSt_MovDEItem = ", mCstIcmsSt)
    Call sqlValor(2, "ValorBCICMSST_MovDEItem = ", mValorBCICMSST)
    Call sqlValor(2, "AliquotaIcmsSt_MovDEItem = ", mAliquotaIcmsSt)
    Call sqlValor(2, "ValorICMSST_MovDEItem = ", mValorICMSST)
    Call sqlTexto(2, "CSTIPI_MovDEItem = ", mCSTIPI)
    Call sqlValor(2, "ValorBcIpi_MovDEItem = ", mValorBcIpi)
    Call sqlValor(2, "AliquotaIpi_MovDEItem = ", mAliquotaIpi)
    Call sqlValor(2, "ValorIPI_MovDEItem = ", mValorIPI)
    Call sqlBoolean(2, "ApuracaoIpiMensal_MovDEItem = ", mApuracaoIpiMensal)
    Call sqlTexto(2, "CodigoEnquadramentoIpi_MovDEItem = ", mCodigoEnquadramentoIpi)
    Call sqlTexto(2, "CSTPIS_MovDEItem = ", mCstPis)
    Call sqlValor(2, "ValorBcPis_MovDEItem = ", mValorBcPis)
    Call sqlValor(2, "QuantidadeBcPis_MovDEItem = ", mQuantidadeBcPis)
    Call sqlValor(2, "AliquotaPis_MovDEItem = ", mAliquotaPis)
    Call sqlValor(2, "ValorPis_MovDEItem = ", mValorPis)
    Call sqlTexto(2, "CSTCOFINS_MovDEItem = ", mCstCofins)
    Call sqlValor(2, "ValorBcCofins_MovDEItem = ", mValorBcCofins)
    Call sqlValor(2, "QuantidadeBcCofins_MovDEItem = ", mQuantidadeBcCofins)
    Call sqlValor(2, "ValorCofins_MovDEItem = ", mValorCofins)
    Call sqlNumero(2, "NumeroTanque_MovDEItem = ", mNumeroTanque)
    Call sqlBoolean(2, "Cancelado_MovDEItem = ", mCancelado)
    Call sqlData(2, "DataEntradaSaida_MovDEItem = ", mDataEntradaSaida)
    Call sqlValor(2, "EncerranteFinal_MovDEItem = ", mEncerranteFinal)
    Call sqlNumero(2, "NumeroBomba_MovDEItem = ", mNumeroBomba)
    Call sqlNumero(2, "NumeroBico_MovDEItem = ", mNumeroBico)
    Call sqlTexto(2, "TipoCombustivel_MovDEItem = ", mTipoCombustivel)
    Call sqlNumero(2, "EtapaConcluida_MovDEItem = ", mEtapaConcluida)
    Call sqlTexto(2, "ProgramaOrigem_MovDEItem = ", mProgramaOrigem)
    Call sqlTexto(2, "Periodo_MovDEItem = ", mPeriodo)
    
    
    gSQL = gSQL & " WHERE IdEstabelecimento_MovDEItem = " & pIdEstabelecimento
    gSQL = gSQL & " AND DataEmissao_MovDEItem = " & preparaData(pDataEmissao)
    gSQL = gSQL & " AND Entrada_MovDEItem = " & preparaBooleano(pEntrada)
    gSQL = gSQL & " AND Saida_MovDEItem = " & preparaBooleano(pSaida)
    gSQL = gSQL & " AND Modelo_MovDEItem = " & preparaTexto(pModelo)
    gSQL = gSQL & " AND Serie_MovDEItem = " & preparaTexto(pSerie)
    gSQL = gSQL & " AND Numero_MovDEItem = " & preparaTexto(pNumero)
    gSQL = gSQL & " AND Ordem_MovDEItem = " & pOrdem
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    '        Call IncluiIntegracaoNuvem("ALTERAR", pEmpresa, pCodigoECF, pNumeroCupom, pData, pOrdem)
    End If
    Exit Function


Err_Alterar:

End Function
Public Function AlterarDesconto(ByVal pIdEstabelecimento As Integer, ByVal pDataEmissao As Date, ByVal pEntrada As Boolean, ByVal pSaida As Boolean, ByVal pModelo As String, ByVal pSerie As String, ByVal pNumero As String, ByVal pTotalCupom As Currency, ByVal pValorDesconto As Currency) As Boolean

Dim lRecordsAffected As Long
Dim xSomaDesconto As Currency
Dim xPercentual As Currency
Dim xValorDesconto As Currency
Dim xOrdem As Integer
Dim xValorBaseCalculoPis As Currency
Dim xValorBaseCalculoCofins As Currency
Dim xValorBaseCalculoIcms As Currency
Dim xValorCofins As Currency
Dim xValorPis As Currency
Dim xValorIcms As Currency

On Error GoTo Err_Alterar

    AlterarDesconto = False
    xSomaDesconto = 0
    xPercentual = Format(pValorDesconto * 100 / pTotalCupom, "0000000000.0000")
    gSQL = "SELECT ValorDesconto_MovDEItem, ValorTotalLiquido_MovDEItem, Ordem_MovDEItem, ValorBcIcms_MovDeItem, ValorIcms_MovDeItem, AliquotaIcms_MovDeItem, ValorBcPis_MovDEItem, ValorBcCofins_MovDEItem, ValorPis_MovDEItem, ValorCofins_MovDEItem, AliquotaPis_MovDEItem, AliquotaCofins_MovDEItem"
    gSQL = gSQL & "  FROM MovimentoDocumentoEletronicoItem"
    gSQL = gSQL & " WHERE IdEstabelecimento_MovDEItem = " & pIdEstabelecimento
    gSQL = gSQL & "   AND DataEmissao_MovDEItem = " & preparaData(pDataEmissao)
    gSQL = gSQL & "   AND Entrada_MovDEItem = " & preparaBooleano(pEntrada)
    gSQL = gSQL & "   AND Saida_MovDEItem = " & preparaBooleano(pSaida)
    gSQL = gSQL & "   AND Modelo_MovDEItem = " & preparaTexto(pModelo)
    gSQL = gSQL & "   AND Serie_MovDEItem = " & preparaTexto(pSerie)
    gSQL = gSQL & "   AND Numero_MovDEItem = " & preparaTexto(pNumero)
    gSQL = gSQL & "   AND Cancelado_MovDEItem = " & preparaBooleano(False)
    gSQL = gSQL & " ORDER BY Ordem_MovDEItem"
    If Localizar(99) Then
        If rsMovDocumentoEletronicoItem.RecordCount > 1 Then
            Do Until rsMovDocumentoEletronicoItem.EOF
                xOrdem = rsMovDocumentoEletronicoItem("Ordem_MovDEItem").Value
                xValorDesconto = Format(rsMovDocumentoEletronicoItem("ValorTotalLiquido_MovDEItem").Value * xPercentual / 100, "0000000000.0000")
                xSomaDesconto = xSomaDesconto + xValorDesconto
                
                If rsMovDocumentoEletronicoItem("ValorPis_MovDEItem").Value > 0 Then
                    xValorBaseCalculoPis = rsMovDocumentoEletronicoItem("ValorTotalLiquido_MovDEItem").Value - xValorDesconto
                    xValorPis = FormatNumber(xValorBaseCalculoPis * (rsMovDocumentoEletronicoItem("AliquotaPis_MovDEItem").Value / 100), 2)
                Else
                    xValorBaseCalculoPis = 0
                    xValorPis = 0
                End If
                
                If rsMovDocumentoEletronicoItem("ValorCofins_MovDEItem").Value > 0 Then
                    xValorBaseCalculoCofins = rsMovDocumentoEletronicoItem("ValorTotalLiquido_MovDEItem").Value - xValorDesconto
                    xValorCofins = FormatNumber(xValorBaseCalculoCofins * (rsMovDocumentoEletronicoItem("AliquotaCofins_MovDEItem").Value / 100), 2)
                Else
                    xValorBaseCalculoCofins = 0
                    xValorCofins = 0
                End If
                
                If rsMovDocumentoEletronicoItem("ValorIcms_MovDEItem").Value > 0 Then
                    xValorBaseCalculoIcms = rsMovDocumentoEletronicoItem("ValorTotalLiquido_MovDEItem").Value - xValorDesconto
                    xValorIcms = FormatNumber(xValorBaseCalculoIcms * (rsMovDocumentoEletronicoItem("AliquotaIcms_MovDEItem").Value / 100), 2)
                Else
                    xValorBaseCalculoIcms = 0
                    xValorIcms = 0
                End If
                
                'Altera
                gSQL = "UPDATE MovimentoDocumentoEletronicoItem SET "
                Call sqlValor(2, "ValorDesconto_MovDEItem = ", xValorDesconto)
                Call sqlValor(2, ", ValorTotalLiquido_MovDEItem = ValorTotalLiquido_MovDEItem - ", xValorDesconto)
                
                If (xValorBaseCalculoPis <> rsMovDocumentoEletronicoItem("ValorBcPis_MovDEItem").Value) Then
                    Call sqlValor(2, ", ValorBcPis_MovDEItem = ", xValorBaseCalculoPis)
                    Call sqlValor(2, ", ValorPis_MovDEItem = ", xValorPis)
                End If
                
                If (xValorBaseCalculoCofins <> rsMovDocumentoEletronicoItem("ValorBcCofins_MovDEItem").Value) Then
                    Call sqlValor(2, ", ValorBcCofins_MovDEItem = ", xValorBaseCalculoCofins)
                    Call sqlValor(2, ", ValorCofins_MovDEItem = ", xValorCofins)
                End If
                
                If (xValorBaseCalculoIcms <> rsMovDocumentoEletronicoItem("ValorBcIcms_MovDEItem").Value) Then
                    Call sqlValor(2, ", ValorBcIcms_MovDEItem = ", xValorBaseCalculoIcms)
                    Call sqlValor(2, ", ValorIcms_MovDEItem = ", xValorIcms)
                End If
                
                gSQL = gSQL & " WHERE IdEstabelecimento_MovDEItem = " & pIdEstabelecimento
                gSQL = gSQL & "   AND DataEmissao_MovDEItem = " & preparaData(pDataEmissao)
                gSQL = gSQL & "   AND Entrada_MovDEItem = " & preparaBooleano(pEntrada)
                gSQL = gSQL & "   AND Saida_MovDEItem = " & preparaBooleano(pSaida)
                gSQL = gSQL & "   AND Modelo_MovDEItem = " & preparaTexto(pModelo)
                gSQL = gSQL & "   AND Serie_MovDEItem = " & preparaTexto(pSerie)
                gSQL = gSQL & "   AND Numero_MovDEItem = " & preparaTexto(pNumero)
                gSQL = gSQL & "   AND Ordem_MovDEItem = " & rsMovDocumentoEletronicoItem("Ordem_MovDEItem").Value
                gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
                If lRecordsAffected > 0 Then
                    AlterarDesconto = True
                End If
                rsMovDocumentoEletronicoItem.MoveNext
            Loop
            If pValorDesconto <> FormatNumber(xSomaDesconto, 2) Then
                rsMovDocumentoEletronicoItem.MoveLast 'volto pro ultimo item e coloco  a diferença nele
                If pValorDesconto < xSomaDesconto Then
                    xValorDesconto = xValorDesconto - (xSomaDesconto - pValorDesconto)
                Else
                    xValorDesconto = xValorDesconto + (pValorDesconto - xSomaDesconto)
                End If
                
                If rsMovDocumentoEletronicoItem("ValorPis_MovDEItem").Value > 0 Then
                    xValorBaseCalculoPis = rsMovDocumentoEletronicoItem("ValorTotalLiquido_MovDEItem").Value - xValorDesconto
                    xValorPis = FormatNumber(xValorBaseCalculoPis * (rsMovDocumentoEletronicoItem("AliquotaPis_MovDEItem").Value / 100), 2)
                Else
                    xValorBaseCalculoPis = 0
                    xValorPis = 0
                End If
                
                If rsMovDocumentoEletronicoItem("ValorCofins_MovDEItem").Value > 0 Then
                    xValorBaseCalculoCofins = rsMovDocumentoEletronicoItem("ValorTotalLiquido_MovDEItem").Value - xValorDesconto
                    xValorCofins = FormatNumber(xValorBaseCalculoCofins * (rsMovDocumentoEletronicoItem("AliquotaCofins_MovDEItem").Value / 100), 2)
                Else
                    xValorBaseCalculoCofins = 0
                    xValorCofins = 0
                End If
                
                If rsMovDocumentoEletronicoItem("ValorIcms_MovDEItem").Value > 0 Then
                    xValorBaseCalculoIcms = rsMovDocumentoEletronicoItem("ValorTotalLiquido_MovDEItem").Value - xValorDesconto
                    xValorIcms = FormatNumber(xValorBaseCalculoIcms * (rsMovDocumentoEletronicoItem("AliquotaIcms_MovDEItem").Value / 100), 2)
                Else
                    xValorBaseCalculoIcms = 0
                    xValorIcms = 0
                End If
                
                gSQL = "UPDATE MovimentoDocumentoEletronicoItem SET "
                Call sqlValor(2, "ValorDesconto_MovDEItem = ", xValorDesconto)
                Call sqlValor(2, ", ValorTotalLiquido_MovDEItem = ValorTotalLiquido_MovDEItem - ", xValorDesconto)
                
                If (xValorBaseCalculoPis <> rsMovDocumentoEletronicoItem("ValorBcPis_MovDEItem").Value) Then
                    Call sqlValor(2, ", ValorBcPis_MovDEItem = ", xValorBaseCalculoPis)
                    Call sqlValor(2, ", ValorPis_MovDEItem = ", xValorPis)
                End If
                
                If (xValorBaseCalculoCofins <> rsMovDocumentoEletronicoItem("ValorBcCofins_MovDEItem").Value) Then
                    Call sqlValor(2, ", ValorBcCofins_MovDEItem = ", xValorBaseCalculoCofins)
                    Call sqlValor(2, ", ValorCofins_MovDEItem = ", xValorCofins)
                End If
                
                If (xValorBaseCalculoIcms <> rsMovDocumentoEletronicoItem("ValorBcIcms_MovDEItem").Value) Then
                    Call sqlValor(2, ", ValorBcIcms_MovDEItem = ", xValorBaseCalculoIcms)
                    Call sqlValor(2, ", ValorIcms_MovDEItem = ", xValorIcms)
                End If

                
                gSQL = gSQL & " WHERE IdEstabelecimento_MovDEItem = " & pIdEstabelecimento
                gSQL = gSQL & "   AND DataEmissao_MovDEItem = " & preparaData(pDataEmissao)
                gSQL = gSQL & "   AND Entrada_MovDEItem = " & preparaBooleano(pEntrada)
                gSQL = gSQL & "   AND Saida_MovDEItem = " & preparaBooleano(pSaida)
                gSQL = gSQL & "   AND Modelo_MovDEItem = " & preparaTexto(pModelo)
                gSQL = gSQL & "   AND Serie_MovDEItem = " & preparaTexto(pSerie)
                gSQL = gSQL & "   AND Numero_MovDEItem = " & preparaTexto(pNumero)
                gSQL = gSQL & "   AND Ordem_MovDEItem = " & rsMovDocumentoEletronicoItem("Ordem_MovDEItem").Value
                gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
                If lRecordsAffected > 0 Then
                    AlterarDesconto = True
                End If
            End If
        Else
        
            If rsMovDocumentoEletronicoItem("ValorPis_MovDEItem").Value > 0 Then
                xValorBaseCalculoPis = rsMovDocumentoEletronicoItem("ValorTotalLiquido_MovDEItem").Value - pValorDesconto
                xValorPis = FormatNumber(xValorBaseCalculoPis * (rsMovDocumentoEletronicoItem("AliquotaPis_MovDEItem").Value / 100), 2)
            Else
                xValorBaseCalculoPis = 0
                xValorPis = 0
            End If
                
            If rsMovDocumentoEletronicoItem("ValorCofins_MovDEItem").Value > 0 Then
                xValorBaseCalculoCofins = rsMovDocumentoEletronicoItem("ValorTotalLiquido_MovDEItem").Value - pValorDesconto
                xValorCofins = FormatNumber(xValorBaseCalculoCofins * (rsMovDocumentoEletronicoItem("AliquotaCofins_MovDEItem").Value / 100), 2)
            Else
                xValorBaseCalculoCofins = 0
                xValorCofins = 0
            End If

            If rsMovDocumentoEletronicoItem("ValorIcms_MovDEItem").Value > 0 Then
                xValorBaseCalculoIcms = rsMovDocumentoEletronicoItem("ValorTotalLiquido_MovDEItem").Value - pValorDesconto
                xValorIcms = FormatNumber(xValorBaseCalculoIcms * (rsMovDocumentoEletronicoItem("AliquotaIcms_MovDEItem").Value / 100), 2)
            Else
                xValorBaseCalculoIcms = 0
                xValorIcms = 0
            End If
        
            gSQL = "UPDATE MovimentoDocumentoEletronicoItem SET "
            Call sqlValor(2, "ValorDesconto_MovDEItem = ", pValorDesconto)
            Call sqlValor(2, ", ValorTotalLiquido_MovDEItem = ValorTotalLiquido_MovDEItem - ", pValorDesconto)
            
            If (xValorBaseCalculoPis <> rsMovDocumentoEletronicoItem("ValorBcPis_MovDEItem").Value) Then
                Call sqlValor(2, ", ValorBcPis_MovDEItem = ", xValorBaseCalculoPis)
                Call sqlValor(2, ", ValorPis_MovDEItem = ", xValorPis)
            End If
                
            If (xValorBaseCalculoCofins <> rsMovDocumentoEletronicoItem("ValorBcCofins_MovDEItem").Value) Then
               Call sqlValor(2, ", ValorBcCofins_MovDEItem = ", xValorBaseCalculoCofins)
               Call sqlValor(2, ", ValorCofins_MovDEItem = ", xValorCofins)
            End If

            If (xValorBaseCalculoIcms <> rsMovDocumentoEletronicoItem("ValorBcIcms_MovDEItem").Value) Then
               Call sqlValor(2, ", ValorBcIcms_MovDEItem = ", xValorBaseCalculoIcms)
               Call sqlValor(2, ", ValorIcms_MovDEItem = ", xValorIcms)
            End If
            
            gSQL = gSQL & " WHERE IdEstabelecimento_MovDEItem = " & pIdEstabelecimento
            gSQL = gSQL & "   AND DataEmissao_MovDEItem = " & preparaData(pDataEmissao)
            gSQL = gSQL & "   AND Entrada_MovDEItem = " & preparaBooleano(pEntrada)
            gSQL = gSQL & "   AND Saida_MovDEItem = " & preparaBooleano(pSaida)
            gSQL = gSQL & "   AND Modelo_MovDEItem = " & preparaTexto(pModelo)
            gSQL = gSQL & "   AND Serie_MovDEItem = " & preparaTexto(pSerie)
            gSQL = gSQL & "   AND Numero_MovDEItem = " & preparaTexto(pNumero)
            gSQL = gSQL & "   AND Ordem_MovDEItem = " & rsMovDocumentoEletronicoItem("Ordem_MovDEItem").Value
            gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
            If lRecordsAffected > 0 Then
                AlterarDesconto = True
            End If
        End If
    End If
    rsMovDocumentoEletronicoItem.Close
    Set rsMovDocumentoEletronicoItem = Nothing
    Exit Function

Err_Alterar:
End Function

Public Function AlterarEtapaConcluida(ByVal pIdEstabelecimento As Integer, ByVal pDataEmissao As Date, ByVal pEntrada As Boolean, ByVal pSaida As Boolean, ByVal pModelo As String, ByVal pSerie As String, ByVal pNumero As String, ByVal pEtapaConcluida As Integer) As Boolean

Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    AlterarEtapaConcluida = False
    gSQL = "UPDATE MovimentoDocumentoEletronicoItem SET "
    Call sqlNumero(2, "EtapaConcluida_MovDEItem = ", pEtapaConcluida)
    gSQL = gSQL & " WHERE IdEstabelecimento_MovDEItem = " & pIdEstabelecimento
    gSQL = gSQL & "   AND DataEmissao_MovDEItem = " & preparaData(pDataEmissao)
    gSQL = gSQL & "   AND Entrada_MovDEItem = " & preparaBooleano(pEntrada)
    gSQL = gSQL & "   AND Saida_MovDEItem = " & preparaBooleano(pSaida)
    gSQL = gSQL & "   AND Modelo_MovDEItem = " & preparaTexto(pModelo)
    gSQL = gSQL & "   AND Serie_MovDEItem = " & preparaTexto(pSerie)
    gSQL = gSQL & "   AND Numero_MovDEItem = " & preparaTexto(pNumero)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlterarEtapaConcluida = True
    End If
    Exit Function

Err_Alterar:
End Function
Public Function AlterarIdCliente(ByVal pIdEstabelecimento As Integer, ByVal pDataEmissao As Date, ByVal pEntrada As Boolean, ByVal pSaida As Boolean, ByVal pModelo As String, ByVal pSerie As String, ByVal pNumero As String, ByVal pIdCliente As Integer) As Boolean

Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    AlterarIdCliente = False
    gSQL = "UPDATE MovimentoDocumentoEletronicoItem SET "
    Call sqlNumero(2, "IdClienteFornecedor_MovDEItem = ", pIdCliente)
    gSQL = gSQL & " WHERE IdEstabelecimento_MovDEItem = " & pIdEstabelecimento
    gSQL = gSQL & "   AND DataEmissao_MovDEItem = " & preparaData(pDataEmissao)
    gSQL = gSQL & "   AND Entrada_MovDEItem = " & preparaBooleano(pEntrada)
    gSQL = gSQL & "   AND Saida_MovDEItem = " & preparaBooleano(pSaida)
    gSQL = gSQL & "   AND Modelo_MovDEItem = " & preparaTexto(pModelo)
    gSQL = gSQL & "   AND Serie_MovDEItem = " & preparaTexto(pSerie)
    gSQL = gSQL & "   AND Numero_MovDEItem = " & preparaTexto(pNumero)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlterarIdCliente = True
    End If
    Exit Function

Err_Alterar:
End Function
Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovDocumentoEletronicoItem = New adodb.Recordset
    rsMovDocumentoEletronicoItem.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovDocumentoEletronicoItem.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovDocumentoEletronicoItem.EOF Then
        AtualizaRecordset = True
    End If
End Function
Private Sub AtribuiValor()

    IdEstabelecimento = rsMovDocumentoEletronicoItem("IdEstabelecimento_MovDEItem").Value
    DataEmissao = rsMovDocumentoEletronicoItem("DataEmissao_MovDEItem").Value
    Entrada = rsMovDocumentoEletronicoItem("Entrada_MovDEItem").Value
    Saida = rsMovDocumentoEletronicoItem("Saida_MovDEItem").Value
    Modelo = rsMovDocumentoEletronicoItem("Modelo_MovDEItem").Value
    Serie = rsMovDocumentoEletronicoItem("Serie_MovDEItem").Value
    Numero = rsMovDocumentoEletronicoItem("Numero_MovDEItem").Value
    Ordem = rsMovDocumentoEletronicoItem("Ordem_MovDEItem").Value
    IdClienteFornecedor = rsMovDocumentoEletronicoItem("IdClienteFornecedor_MovDEItem").Value
    MovimentacaoFisica = rsMovDocumentoEletronicoItem("MovimentacaoFisica_MovDEItem").Value
    CFOP = rsMovDocumentoEletronicoItem("CFOP_MovDEItem").Value
    IdProduto = rsMovDocumentoEletronicoItem("IdProduto_MovDEItem").Value
    ValorUnitario = rsMovDocumentoEletronicoItem("ValorUnitario_MovDEItem").Value
    Quantidade = rsMovDocumentoEletronicoItem("Quantidade_MovDEItem").Value
    ValorDesconto = rsMovDocumentoEletronicoItem("ValorDesconto_MovDEItem").Value
    ValorTotalLiquido = rsMovDocumentoEletronicoItem("ValorTotalLiquido_MovDEItem").Value
    CSTICMS = rsMovDocumentoEletronicoItem("CSTICMS_MovDEItem").Value
    ValorBCICMS = rsMovDocumentoEletronicoItem("ValorBCICMS_MovDEItem").Value
    AliquotaICMS = rsMovDocumentoEletronicoItem("AliquotaICMS_MovDEItem").Value
    ValorICMS = rsMovDocumentoEletronicoItem("ValorICMS_MovDEItem").Value
    CstIcmsSt = rsMovDocumentoEletronicoItem("CstIcmsSt_MovDEItem").Value
    ValorBCICMSST = rsMovDocumentoEletronicoItem("ValorBCICMSST_MovDEItem").Value
    AliquotaIcmsSt = rsMovDocumentoEletronicoItem("AliquotaIcmsSt_MovDEItem").Value
    ValorICMSST = rsMovDocumentoEletronicoItem("ValorICMSST_MovDEItem").Value
    CSTIPI = rsMovDocumentoEletronicoItem("CSTIPI_MovDEItem").Value
    ValorBcIpi = rsMovDocumentoEletronicoItem("ValorBcIpi_MovDEItem").Value
    AliquotaIpi = rsMovDocumentoEletronicoItem("AliquotaIpi_MovDEItem").Value
    ValorIPI = rsMovDocumentoEletronicoItem("ValorIPI_MovDEItem").Value
    ApuracaoIpiMensal = rsMovDocumentoEletronicoItem("ApuracaoIpiMensal_MovDEItem").Value
    CodigoEnquadramentoIpi = rsMovDocumentoEletronicoItem("CodigoEnquadramentoIpi_MovDEItem").Value
    CstPis = rsMovDocumentoEletronicoItem("CSTPIS_MovDEItem").Value
    ValorBcPis = rsMovDocumentoEletronicoItem("ValorBcPis_MovDEItem").Value
    QuantidadeBcPis = rsMovDocumentoEletronicoItem("QuantidadeBcPis_MovDEItem").Value
    AliquotaPis = rsMovDocumentoEletronicoItem("AliquotaPis_MovDEItem").Value
    ValorPis = rsMovDocumentoEletronicoItem("ValorPis_MovDEItem").Value
    CstCofins = rsMovDocumentoEletronicoItem("CSTCOFINS_MovDEItem").Value
    ValorBcCofins = rsMovDocumentoEletronicoItem("ValorBcCofins_MovDEItem").Value
    QuantidadeBcCofins = rsMovDocumentoEletronicoItem("QuantidadeBcCofins_MovDEItem").Value
    ValorCofins = rsMovDocumentoEletronicoItem("ValorCofins_MovDEItem").Value
    NumeroTanque = rsMovDocumentoEletronicoItem("NumeroTanque_MovDEItem").Value
    Cancelado = rsMovDocumentoEletronicoItem("Cancelado_MovDEItem").Value
    DataEntradaSaida = rsMovDocumentoEletronicoItem("DataEntradaSaida_MovDEItem").Value
    EncerranteFinal = rsMovDocumentoEletronicoItem("EncerranteFinal_MovDEItem").Value
    NumeroBomba = rsMovDocumentoEletronicoItem("NumeroBomba_MovDEItem").Value
    NumeroBico = rsMovDocumentoEletronicoItem("NumeroBico_MovDEItem").Value
    TipoCombustivel = rsMovDocumentoEletronicoItem("TipoCombustivel_MovDEItem").Value
    EtapaConcluida = rsMovDocumentoEletronicoItem("EtapaConcluida_MovDEItem").Value
    ProgramaOrigem = rsMovDocumentoEletronicoItem("ProgramaOrigem_MovDEItem").Value
    Periodo = rsMovDocumentoEletronicoItem("Periodo_MovDEItem").Value
End Sub
'exclui todos itens do documento abandonado
Public Function ExcluirItensDocumento(ByVal pIdEstabelecimento As Integer, ByVal pDataEmissao As Date, ByVal pEntrada As Boolean, ByVal pSaida As Boolean, ByVal pModelo As String, ByVal pSerie As String, ByVal pNumero As String) As Boolean
    
    Dim lRecordsAffected As Long
    Dim xCondicao As String

On Error GoTo Err_Excluir

    ExcluirItensDocumento = False
    gSQL = "DELETE FROM MovimentoDocumentoEletronicoItem"
    xCondicao = " WHERE IdEstabelecimento_MovDEItem = " & pIdEstabelecimento
    xCondicao = xCondicao & " AND DataEmissao_MovDEItem = " & preparaData(pDataEmissao)
    xCondicao = xCondicao & " AND Entrada_MovDEItem = " & preparaBooleano(pEntrada)
    xCondicao = xCondicao & " AND Saida_MovDEItem = " & preparaBooleano(pSaida)
    xCondicao = xCondicao & " AND Modelo_MovDEItem = " & preparaTexto(pModelo)
    xCondicao = xCondicao & " AND Serie_MovDEItem = " & preparaTexto(pSerie)
    xCondicao = xCondicao & " AND Numero_MovDEItem = " & preparaTexto(pNumero)
    
    gSQL = gSQL & xCondicao
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    
    If lRecordsAffected > 0 Then
        ExcluirItensDocumento = True
    End If
    Exit Function

Err_Excluir:

End Function

Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

Public Function LocalizarCodigo(ByVal pIdEstabelecimento As Integer, ByVal pDataEmissao As Date, ByVal pEntrada As Boolean, ByVal pSaida As Boolean, ByVal pModelo As String, ByVal pSerie As String, ByVal pNumero As String, ByVal pOrdem As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE IdEstabelecimento_MovDEItem = " & pIdEstabelecimento
    xCondicao = xCondicao & " AND DataEmissao_MovDEItem = " & preparaData(pDataEmissao)
    xCondicao = xCondicao & " AND Entrada_MovDEItem = " & preparaBooleano(pEntrada)
    xCondicao = xCondicao & " AND Saida_MovDEItem = " & preparaBooleano(pSaida)
    xCondicao = xCondicao & " AND Modelo_MovDEItem = " & preparaTexto(pModelo)
    xCondicao = xCondicao & " AND Serie_MovDEItem = " & preparaTexto(pSerie)
    xCondicao = xCondicao & " AND Numero_MovDEItem = " & preparaTexto(pNumero)
    xCondicao = xCondicao & " AND Ordem_MovDEItem = " & pOrdem
    
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsMovDocumentoEletronicoItem.Close
    Set rsMovDocumentoEletronicoItem = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
Public Function LocalizarProximoDestaOrdem(ByVal pEmpresa As Integer, ByVal pDataEmissao As Date, ByVal pEntrada As Boolean, ByVal pSaida As Boolean, ByVal pModelo As String, ByVal pSerie As String, ByVal pNumero As String, ByVal pOrdem As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximoDestaOrdem = False
    xCondicao = " WHERE IdEstabelecimento_MovDEItem = " & pEmpresa
    xCondicao = xCondicao & " AND DataEmissao_MovDEItem = " & preparaData(pDataEmissao)
    xCondicao = xCondicao & " AND Entrada_MovDEItem = " & preparaBooleano(pEntrada)
    xCondicao = xCondicao & " AND Saida_MovDEItem = " & preparaBooleano(pSaida)
    xCondicao = xCondicao & " AND Modelo_MovDEItem = " & preparaTexto(pModelo)
    xCondicao = xCondicao & " AND Serie_MovDEItem = " & preparaTexto(pSerie)
    xCondicao = xCondicao & " AND Numero_MovDEItem = " & preparaTexto(pNumero)
    xCondicao = xCondicao & " AND Ordem_MovDEItem > " & pOrdem
    gSQL = PreparaSQL(xCondicao, "ORDER BY Ordem_MovDEItem ASC")
    If Localizar(1) Then
        LocalizarProximoDestaOrdem = True
        AtribuiValor
    End If
    rsMovDocumentoEletronicoItem.Close
    Set rsMovDocumentoEletronicoItem = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function



Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    Dim xSQL As String
           
    xSQL = "SELECT IdEstabelecimento_MovDEItem,DataEmissao_MovDEItem,Entrada_MovDEItem,Saida_MovDEItem,Modelo_MovDEItem,Serie_MovDEItem,Numero_MovDEItem"
    xSQL = xSQL & ",Ordem_MovDEItem,IdClienteFornecedor_MovDEItem,MovimentacaoFisica_MovDEItem,CFOP_MovDEItem,IdProduto_MovDEItem,ValorUnitario_MovDEItem"
    xSQL = xSQL & ",Quantidade_MovDEItem,ValorDesconto_MovDEItem,ValorTotalLiquido_MovDEItem,CstIcms_MovDEItem,ValorBcIcms_MovDEItem,AliquotaIcms_MovDEItem"
    xSQL = xSQL & ",ValorIcms_MovDEItem,CstIcmsSt_MovDEItem,ValorBcIcmsSt_MovDEItem,AliquotaIcmsSt_MovDEItem,ValorIcmsSt_MovDEItem,CstIpi_MovDEItem"
    xSQL = xSQL & ",ValorBcIpi_MovDEItem,AliquotaIpi_MovDEItem,ValorIpi_MovDEItem,ApuracaoIpiMensal_MovDEItem,CodigoEnquadramentoIpi_MovDEItem"
    xSQL = xSQL & ",CstPis_MovDEItem,ValorBcPis_MovDEItem,QuantidadeBcPis_MovDEItem,AliquotaPis_MovDEItem,ValorPis_MovDEItem,CstCofins_MovDEItem"
    xSQL = xSQL & ",ValorBcCofins_MovDEItem,QuantidadeBcCofins_MovDEItem,AliquotaCofins_MovDEItem,ValorCofins_MovDEItem,NumeroTanque_MovDEItem"
    xSQL = xSQL & ",Cancelado_MovDEItem,DataEntradaSaida_MovDEItem, EncerranteFinal_MovDEItem, NumeroBomba_MovDEItem "
    xSQL = xSQL & ",NumeroBico_MovDEItem, TipoCombustivel_MovDEItem, EtapaConcluida_MovDEItem, ProgramaOrigem_MovDEItem, Periodo_MovDEItem FROM MovimentoDocumentoEletronicoItem"
    
    PreparaSQL = Trim(xSQL & " " & xCondicao & " " & xTipoOrdem)
End Function

