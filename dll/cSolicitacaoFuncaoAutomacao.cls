VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSolicitacaoFuncaoAutomacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mNSU As Long
Private mNumeroControleSolicitacao As Long
Private mDataSolicitacao As Date
Private mHoraSolicitacao As Date
Private mTipoOperacao As String
Private mCodigoEmpresa As Long
Private mIPComputadorAC As String
Private mIPInternetAC As String
Private mSegurancaEstabelecimento As String
Private mCodigoUsuario As Long
Private mVersaoAC As String
Private mVersaoHost As String
Private mTexto As String
Private mHoraAnalise As Date
Private mHoraAprovacao As Date
Private mHoraCancelamentoHost As Date
Private mHoraConfirmacaoAC As Date
Private mHoraCancelamentoAC As Date
Private mMensagem As String

Private rsSolicitacaoFuncaoAutomacao As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get NSU() As Long
    NSU = mNSU
End Property
Public Property Let NSU(ByVal Valor As Long)
    mNSU = Valor
End Property


Public Property Get NumeroControleSolicitacao() As Long
    NumeroControleSolicitacao = mNumeroControleSolicitacao
End Property
Public Property Let NumeroControleSolicitacao(ByVal Valor As Long)
    mNumeroControleSolicitacao = Valor
End Property


Public Property Get DataSolicitacao() As Date
    DataSolicitacao = mDataSolicitacao
End Property
Public Property Let DataSolicitacao(ByVal Valor As Date)
    mDataSolicitacao = Valor
End Property


Public Property Get HoraSolicitacao() As Date
    HoraSolicitacao = mHoraSolicitacao
End Property
Public Property Let HoraSolicitacao(ByVal Valor As Date)
    mHoraSolicitacao = Valor
End Property


Public Property Get TipoOperacao() As String
    TipoOperacao = mTipoOperacao
End Property
Public Property Let TipoOperacao(ByVal Valor As String)
    mTipoOperacao = Valor
End Property


Public Property Get CodigoEmpresa() As Long
    CodigoEmpresa = mCodigoEmpresa
End Property
Public Property Let CodigoEmpresa(ByVal Valor As Long)
    mCodigoEmpresa = Valor
End Property


Public Property Get IPComputadorAC() As String
    IPComputadorAC = mIPComputadorAC
End Property
Public Property Let IPComputadorAC(ByVal Valor As String)
    mIPComputadorAC = Valor
End Property


Public Property Get IPInternetAC() As String
    IPInternetAC = mIPInternetAC
End Property
Public Property Let IPInternetAC(ByVal Valor As String)
    mIPInternetAC = Valor
End Property


Public Property Get SegurancaEstabelecimento() As String
    SegurancaEstabelecimento = mSegurancaEstabelecimento
End Property
Public Property Let SegurancaEstabelecimento(ByVal Valor As String)
    mSegurancaEstabelecimento = Valor
End Property


Public Property Get CodigoUsuario() As Long
    CodigoUsuario = mCodigoUsuario
End Property
Public Property Let CodigoUsuario(ByVal Valor As Long)
    mCodigoUsuario = Valor
End Property


Public Property Get VersaoAC() As String
    VersaoAC = mVersaoAC
End Property
Public Property Let VersaoAC(ByVal Valor As String)
    mVersaoAC = Valor
End Property


Public Property Get VersaoHost() As String
    VersaoHost = mVersaoHost
End Property
Public Property Let VersaoHost(ByVal Valor As String)
    mVersaoHost = Valor
End Property


Public Property Get Texto() As String
    Texto = mTexto
End Property
Public Property Let Texto(ByVal Valor As String)
    mTexto = Valor
End Property


Public Property Get HoraAnalise() As Date
    HoraAnalise = mHoraAnalise
End Property
Public Property Let HoraAnalise(ByVal Valor As Date)
    mHoraAnalise = Valor
End Property


Public Property Get HoraAprovacao() As Date
    HoraAprovacao = mHoraAprovacao
End Property
Public Property Let HoraAprovacao(ByVal Valor As Date)
    mHoraAprovacao = Valor
End Property


Public Property Get HoraCancelamentoHost() As Date
    HoraCancelamentoHost = mHoraCancelamentoHost
End Property
Public Property Let HoraCancelamentoHost(ByVal Valor As Date)
    mHoraCancelamentoHost = Valor
End Property


Public Property Get HoraConfirmacaoAC() As Date
    HoraConfirmacaoAC = mHoraConfirmacaoAC
End Property
Public Property Let HoraConfirmacaoAC(ByVal Valor As Date)
    mHoraConfirmacaoAC = Valor
End Property


Public Property Get HoraCancelamentoAC() As Date
    HoraCancelamentoAC = mHoraCancelamentoAC
End Property
Public Property Let HoraCancelamentoAC(ByVal Valor As Date)
    mHoraCancelamentoAC = Valor
End Property


Public Property Get Mensagem() As String
    Mensagem = mMensagem
End Property
Public Property Let Mensagem(ByVal Valor As String)
    mMensagem = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarNSU(ByVal pNSU As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarNSU = False
    xCondicao = " WHERE NSU = " & pNSU
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) = True Then
        LocalizarNSU = True
        AtribuiValor
    End If
    rsSolicitacaoFuncaoAutomacao.Close
    Set rsSolicitacaoFuncaoAutomacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNCS(ByVal pNumeroControleSolicitacao As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarNCS = False
    xCondicao = " WHERE [Numero de Controle de Solicitacao] = " & pNumeroControleSolicitacao
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        LocalizarNCS = True
        AtribuiValor
    End If
    rsSolicitacaoFuncaoAutomacao.Close
    Set rsSolicitacaoFuncaoAutomacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnalise() As Long

On Error GoTo trata_erro
    
    LocalizarAnalise = 0
    gSQL = "SELECT NSU"
    gSQL = gSQL & "  FROM SolicitacaoFuncaoAutomacao"
    gSQL = gSQL & " WHERE [Hora do Cancelamento AC] IS NULL"
    gSQL = gSQL & "   AND [Hora da Confirmacao AC] IS NULL"
    gSQL = gSQL & "   AND [Hora Cancelamento Host] IS NULL"
    gSQL = gSQL & "   AND [Hora da Analise] IS NULL"
    gSQL = gSQL & "   AND [Hora da Aprovacao] IS NULL"
    gSQL = gSQL & " ORDER BY NSU ASC"
    If Localizar(1) Then
        LocalizarAnalise = rsSolicitacaoFuncaoAutomacao("NSU").Value
    End If
    rsSolicitacaoFuncaoAutomacao.Close
    Set rsSolicitacaoFuncaoAutomacao = Nothing
    Exit Function

trata_erro:
End Function
Public Function LocalizarAnaliseEOuAprovacao() As Boolean

On Error GoTo trata_erro
    
    LocalizarAnaliseEOuAprovacao = False
    gSQL = "SELECT NSU"
    gSQL = gSQL & "  FROM SolicitacaoFuncaoAutomacao"
    gSQL = gSQL & " WHERE [Hora do Cancelamento AC] IS NULL"
    gSQL = gSQL & "   AND [Hora da Confirmacao AC] IS NULL"
    gSQL = gSQL & "   AND [Hora Cancelamento Host] IS NULL"
    gSQL = gSQL & "   AND ([Hora da Analise] IS NULL"
    gSQL = gSQL & "    OR [Hora da Aprovacao] IS NULL)"
    gSQL = gSQL & " ORDER BY NSU ASC"
    LocalizarAnaliseEOuAprovacao = Localizar(1)
    Exit Function

trata_erro:
End Function


Public Function VerificaSeEstaAprovado(ByVal pNSU As Long) As Boolean

On Error GoTo trata_erro
    
    VerificaSeEstaAprovado = False
    gSQL = "SELECT NSU"
    gSQL = gSQL & "  FROM SolicitacaoFuncaoAutomacao"
    gSQL = gSQL & " WHERE NSU = " & pNSU
    gSQL = gSQL & "   AND NOT [Hora da Aprovacao] IS NULL"
    VerificaSeEstaAprovado = Localizar(1)
    Exit Function

trata_erro:
End Function


Public Function VerificaSeEstaEmAnalise(ByVal pNSU As Long) As Boolean

On Error GoTo trata_erro
    
    VerificaSeEstaEmAnalise = False
    gSQL = "SELECT NSU"
    gSQL = gSQL & "  FROM SolicitacaoFuncaoAutomacao"
    gSQL = gSQL & " WHERE NSU = " & pNSU
    gSQL = gSQL & "   AND NOT [Hora da Analise] IS NULL"
    VerificaSeEstaEmAnalise = Localizar(1)
    Exit Function

trata_erro:
End Function


Public Function VerificaSeEstaCanceladoHost(ByVal pNSU As Long) As Boolean

On Error GoTo trata_erro
    
    VerificaSeEstaCanceladoHost = False
    gSQL = "SELECT NSU"
    gSQL = gSQL & "  FROM SolicitacaoFuncaoAutomacao"
    gSQL = gSQL & " WHERE NSU = " & pNSU
    gSQL = gSQL & "   AND NOT [Hora Cancelamento Host] IS NULL"
    VerificaSeEstaCanceladoHost = Localizar(1)
    Exit Function

trata_erro:
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE NSU < " & mNSU
    gSQL = PreparaSQL(xCondicao, "ORDER BY NSU DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsSolicitacaoFuncaoAutomacao.Close
    Set rsSolicitacaoFuncaoAutomacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE [Codigo da Empresa] = " & mCodigoEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY NSU ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsSolicitacaoFuncaoAutomacao.Close
    Set rsSolicitacaoFuncaoAutomacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE NSU > " & mNSU
    gSQL = PreparaSQL(xCondicao, "ORDER BY NSU ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsSolicitacaoFuncaoAutomacao.Close
    Set rsSolicitacaoFuncaoAutomacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY NSU DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsSolicitacaoFuncaoAutomacao.Close
    Set rsSolicitacaoFuncaoAutomacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function ProximoRegistro() As Long
    'Pega Numero do Ultimo movimento
    ProximoRegistro = 1
    gSQL = "SELECT NSU"
    gSQL = gSQL & "  FROM SolicitacaoFuncaoAutomacao"
    gSQL = gSQL & " ORDER BY NSU DESC"
    If Localizar(1) Then
        ProximoRegistro = rsSolicitacaoFuncaoAutomacao("NSU").Value + 1
    End If
    rsSolicitacaoFuncaoAutomacao.Close
    Set rsSolicitacaoFuncaoAutomacao = Nothing
End Function


Public Function Incluir() As Boolean
Dim xRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    mNSU = ProximoRegistro()
    gSQL = "INSERT INTO SolicitacaoFuncaoAutomacao ( NSU, [Numero de Controle de Solicitacao], [Data da Solicitacao], [Hora da Solicitacao], [Tipo de Operacao], "
    gSQL = gSQL & "[Codigo da Empresa], [IP Computador AC], [IP Internet AC], [Seguranca do Estabelecimento], [Codigo do Usuario], "
    gSQL = gSQL & "[Versao AC], [Versao Host], Texto, [Hora da Analise], [Hora da Aprovacao], "
    gSQL = gSQL & "[Hora Cancelamento Host], [Hora da Confirmacao AC], [Hora do Cancelamento AC], Mensagem ) VALUES ( "
    Call sqlNumero(1, mNSU, ", ")
    Call sqlNumero(1, mNumeroControleSolicitacao, ", ")
    Call sqlData(1, mDataSolicitacao, ", ")
    Call sqlHora(1, mHoraSolicitacao, ", ")
    Call sqlTexto(1, mTipoOperacao, ", ")
    Call sqlNumero(1, mCodigoEmpresa, ", ")
    Call sqlTexto(1, mIPComputadorAC, ", ")
    Call sqlTexto(1, mIPInternetAC, ", ")
    Call sqlTexto(1, mSegurancaEstabelecimento, ", ")
    Call sqlNumero(1, mCodigoUsuario, ", ")
    Call sqlTexto(1, mVersaoAC, ", ")
    Call sqlTexto(1, mVersaoHost, ", ")
    Call sqlTexto(1, mTexto, ", ")
    Call sqlHora(1, mHoraAnalise, ", ")
    Call sqlHora(1, mHoraAprovacao, ", ")
    Call sqlHora(1, mHoraCancelamentoHost, ", ")
    Call sqlHora(1, mHoraConfirmacaoAC, ", ")
    Call sqlHora(1, mHoraCancelamentoAC, ", ")
    Call sqlTexto(1, mMensagem, " )")
    gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
    If xRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pNSU As Long) As Boolean
Dim xRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE SolicitacaoFuncaoAutomacao SET "
    Call sqlNumero(2, "NSU = ", mNSU)
    Call sqlNumero(2, ", [Numero de Controle de Solicitacao] = ", mNumeroControleSolicitacao)
    Call sqlData(2, ", [Data da Solicitacao] = ", mDataSolicitacao)
    Call sqlHora(2, ", [Hora da Solicitacao] = ", mHoraSolicitacao)
    Call sqlTexto(2, ", [Tipo de Operacao] = ", mTipoOperacao)
    Call sqlNumero(2, ", [Codigo da Empresa] = ", mCodigoEmpresa)
    Call sqlTexto(2, ", [IP Computador AC] = ", mIPComputadorAC)
    Call sqlTexto(2, ", [IP Internet AC] = ", mIPInternetAC)
    Call sqlTexto(2, ", [Seguranca do Estabelecimento] = ", mSegurancaEstabelecimento)
    Call sqlNumero(2, ", [Codigo do Usuario] = ", mCodigoUsuario)
    Call sqlTexto(2, ", [Versao AC] = ", mVersaoAC)
    Call sqlTexto(2, ", [Versao Host] = ", mVersaoHost)
    Call sqlTexto(2, ", Texto = ", mTexto)
    Call sqlHora(2, ", [Hora da Analise] = ", mHoraAnalise)
    Call sqlHora(2, ", [Hora da Aprovacao] = ", mHoraAprovacao)
    Call sqlHora(2, ", [Hora Cancelamento Host] = ", mHoraCancelamentoHost)
    Call sqlHora(2, ", [Hora da Confirmacao AC] = ", mHoraConfirmacaoAC)
    Call sqlHora(2, ", [Hora do Cancelamento AC] = ", mHoraCancelamentoAC)
    Call sqlTexto(2, ", Mensagem = ", mMensagem)
    gSQL = gSQL & " WHERE NSU = " & pNSU
    gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
    If xRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function BuscaMensagem(ByVal pNSU As Long) As String
    BuscaMensagem = ""
    gSQL = "SELECT Mensagem"
    gSQL = gSQL & "  FROM SolicitacaoFuncaoAutomacao"
    gSQL = gSQL & " WHERE NSU = " & pNSU
    If Localizar(1) Then
        BuscaMensagem = rsSolicitacaoFuncaoAutomacao("Mensagem").Value
    End If
    rsSolicitacaoFuncaoAutomacao.Close
    Set rsSolicitacaoFuncaoAutomacao = Nothing
End Function


Public Function DefineHoraAnalise(ByVal pNSU As Long, ByVal pHoraAnalise As Date, ByVal pVersaoHost As String) As Boolean
Dim xRecordsAffected As Long

On Error GoTo TrataErro
    
    DefineHoraAnalise = False
    gSQL = "UPDATE SolicitacaoFuncaoAutomacao SET "
    Call sqlHora(2, "[Hora da Analise] = ", pHoraAnalise)
    Call sqlTexto(2, ", [Versao Host] = ", Mid(pVersaoHost, 1, 10))
    gSQL = gSQL & " WHERE NSU = " & pNSU
    gSQL = gSQL & "   AND [Hora da Analise] IS NULL"
    gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
    If xRecordsAffected > 0 Then
        DefineHoraAnalise = True
    End If
    Exit Function

TrataErro:
End Function
Public Function DefineHoraAprovacao(ByVal pNSU As Long, ByVal pHoraAprovacao As Date, ByVal pVersaoHost As String, ByVal pMensagem As String) As Boolean
Dim xRecordsAffected As Long

On Error GoTo TrataErro
    
    DefineHoraAprovacao = False
    gSQL = "UPDATE SolicitacaoFuncaoAutomacao SET "
    Call sqlHora(2, "[Hora da Aprovacao] = ", pHoraAprovacao)
    Call sqlTexto(2, ", [Versao Host] = ", Mid(pVersaoHost, 1, 10))
    Call sqlTexto(2, ", Mensagem = ", Mid(pMensagem, 1, 50))
    gSQL = gSQL & " WHERE NSU = " & pNSU
    gSQL = gSQL & "   AND [Hora da Aprovacao] IS NULL"
    gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
    If xRecordsAffected > 0 Then
        DefineHoraAprovacao = True
    End If
    Exit Function

TrataErro:
End Function
Public Function DefineHoraCancelamentoAC(ByVal pNSU As Long, ByVal pHoraCancelamentoAC As Date) As Boolean
Dim xRecordsAffected As Long
    
On Error GoTo TrataErro
    
    DefineHoraCancelamentoAC = False
    gSQL = "UPDATE SolicitacaoFuncaoAutomacao SET "
    Call sqlHora(2, "[Hora do Cancelamento AC] = ", pHoraCancelamentoAC)
    gSQL = gSQL & " WHERE NSU = " & pNSU
    'gSQL = gSQL & "   AND [Hora do Cancelamento AC] IS NULL"
    gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
    If xRecordsAffected > 0 Then
        DefineHoraCancelamentoAC = True
    End If
    Exit Function

TrataErro:
End Function
Public Function DefineHoraCancelamentoHost(ByVal pNSU As Long, ByVal pHoraCancelamentoHost As Date, ByVal pVersaoHost As String, ByVal pMensagem As String) As Boolean
Dim xRecordsAffected As Long
    
On Error GoTo TrataErro
    
    DefineHoraCancelamentoHost = False
    gSQL = "UPDATE SolicitacaoFuncaoAutomacao SET "
    Call sqlHora(2, "[Hora Cancelamento Host] = ", pHoraCancelamentoHost)
    Call sqlTexto(2, ", [Versao Host] = ", Mid(pVersaoHost, 1, 10))
    Call sqlTexto(2, ", Mensagem = ", Mid(pMensagem, 1, 50))
    gSQL = gSQL & " WHERE NSU = " & pNSU
    gSQL = gSQL & "   AND [Hora Cancelamento Host] IS NULL"
    gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
    If xRecordsAffected > 0 Then
        DefineHoraCancelamentoHost = True
    End If
    Exit Function

TrataErro:
End Function
Public Function DefineHoraConfirmacaoAC(ByVal pNSU As Long, ByVal pHoraConfirmacaoAC As Date) As Boolean
Dim xRecordsAffected As Long
    
On Error GoTo TrataErro
    
    DefineHoraConfirmacaoAC = False
    gSQL = "UPDATE SolicitacaoFuncaoAutomacao SET "
    Call sqlHora(2, "[Hora da Confirmacao AC] = ", pHoraConfirmacaoAC)
    gSQL = gSQL & " WHERE NSU = " & pNSU
    gSQL = gSQL & "   AND [Hora da Confirmacao AC] IS NULL"
    gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
    If xRecordsAffected > 0 Then
        DefineHoraConfirmacaoAC = True
    End If
    Exit Function

TrataErro:
End Function
Public Function DefineMensagem(ByVal pNSU As Long, ByVal pMensagem As String) As Boolean
Dim xRecordsAffected As Long
    
On Error GoTo TrataErro
    
    DefineMensagem = False
    gSQL = "UPDATE SolicitacaoFuncaoAutomacao SET "
    Call sqlTexto(2, "Mensagem = ", pMensagem)
    gSQL = gSQL & " WHERE NSU = " & pNSU
    gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
    If xRecordsAffected > 0 Then
        DefineMensagem = True
    End If
    Exit Function

TrataErro:
End Function


Public Function Excluir(ByVal pNSU As Long) As Boolean
Dim xRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM SolicitacaoFuncaoAutomacao"
    gSQL = gSQL & " WHERE NSU = " & pNSU
    gConn.Execute gSQL, xRecordsAffected, adCmdText + adExecuteNoRecords
    If xRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT NSU, [Numero de Controle de Solicitacao], [Data da Solicitacao], [Hora da Solicitacao], [Tipo de Operacao], [Codigo da Empresa], [IP Computador AC], [IP Internet AC], [Seguranca do Estabelecimento], [Codigo do Usuario], [Versao AC], [Versao Host], Texto, [Hora da Analise], [Hora da Aprovacao], [Hora Cancelamento Host], [Hora da Confirmacao AC], [Hora do Cancelamento AC], Mensagem FROM SolicitacaoFuncaoAutomacao" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsSolicitacaoFuncaoAutomacao = New adodb.Recordset
    rsSolicitacaoFuncaoAutomacao.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsSolicitacaoFuncaoAutomacao.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsSolicitacaoFuncaoAutomacao.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    NSU = rsSolicitacaoFuncaoAutomacao("NSU").Value
    NumeroControleSolicitacao = rsSolicitacaoFuncaoAutomacao("Numero de Controle de Solicitacao").Value
    DataSolicitacao = rsSolicitacaoFuncaoAutomacao("Data da Solicitacao").Value
    HoraSolicitacao = rsSolicitacaoFuncaoAutomacao("Hora da Solicitacao").Value
    TipoOperacao = rsSolicitacaoFuncaoAutomacao("Tipo de Operacao").Value
    CodigoEmpresa = rsSolicitacaoFuncaoAutomacao("Codigo da Empresa").Value
    IPComputadorAC = rsSolicitacaoFuncaoAutomacao("IP Computador AC").Value
    IPInternetAC = rsSolicitacaoFuncaoAutomacao("IP Internet AC").Value
    SegurancaEstabelecimento = rsSolicitacaoFuncaoAutomacao("Seguranca do Estabelecimento").Value
    CodigoUsuario = rsSolicitacaoFuncaoAutomacao("Codigo do Usuario").Value
    VersaoAC = rsSolicitacaoFuncaoAutomacao("Versao AC").Value
    VersaoHost = rsSolicitacaoFuncaoAutomacao("Versao Host").Value
    Texto = rsSolicitacaoFuncaoAutomacao("Texto").Value
    If IsNull(rsSolicitacaoFuncaoAutomacao("Hora da Analise").Value) Then
        HoraAnalise = CDate("00:00:00")
    Else
        HoraAnalise = rsSolicitacaoFuncaoAutomacao("Hora da Analise").Value
    End If
    If IsNull(rsSolicitacaoFuncaoAutomacao("Hora da Aprovacao").Value) Then
        HoraAprovacao = CDate("00:00:00")
    Else
        HoraAprovacao = rsSolicitacaoFuncaoAutomacao("Hora da Aprovacao").Value
    End If
    If IsNull(rsSolicitacaoFuncaoAutomacao("Hora Cancelamento Host").Value) Then
        HoraCancelamentoHost = CDate("00:00:00")
    Else
        HoraCancelamentoHost = rsSolicitacaoFuncaoAutomacao("Hora Cancelamento Host").Value
    End If
    If IsNull(rsSolicitacaoFuncaoAutomacao("Hora da Confirmacao AC").Value) Then
        HoraConfirmacaoAC = CDate("00:00:00")
    Else
        HoraConfirmacaoAC = rsSolicitacaoFuncaoAutomacao("Hora da Confirmacao AC").Value
    End If
    If IsNull(rsSolicitacaoFuncaoAutomacao("Hora do Cancelamento AC").Value) Then
        HoraCancelamentoAC = CDate("00:00:00")
    Else
        HoraCancelamentoAC = rsSolicitacaoFuncaoAutomacao("Hora do Cancelamento AC").Value
    End If
    Mensagem = rsSolicitacaoFuncaoAutomacao("Mensagem").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

