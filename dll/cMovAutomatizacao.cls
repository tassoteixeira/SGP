VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovAutomatizacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Private mNSU_MovAutomatizacao As Long
    Private mIdEstabelecimento_MovAutomatizacao As Integer
    Private mDataSolicitacao_MovAutomatizacao As Date
    Private mOperacao_MovAutomatizacao As String
    Private mDados_MovAutomatizacao As String
    Private mIPComputadorAC_MovAutomatizacao As String
    Private mIPInternetAC_MovAutomatizacao As String
    Private mVersaoAC_MovAutomatizacao As String
    Private mVersaoHost_MovAutomatizacao As String
    Private mHoraAnalise_MovAutomatizacao As Date
    Private mEtapaAndamento_MovAutomatizacao As Integer
    Private mDescricaoAndamento_MovAutomatizacao As String
    Private mHoraImportacaoDados_MovAutomatizacao As Date
    Private mHoraImportacaoComplementar_MovAutomatizacao As Date
    Private mHoraConciliacaoDados_MovAutomatizacao As Date
    Private mHoraCancelamentoHost_MovAutomatizacao As Date
    Private mHoraConclusao_MovAutomatizacao As Date
    Private mMensagemRetorno_MovAutomatizacao As String
    Private mObservacao_MovAutomatizacao As String

    Private rsAutomatizacao As New adodb.Recordset
    
    Public Property Let NSU_MovAutomatizacao(ByVal Valor As Long)
        mNSU_MovAutomatizacao = Valor
    End Property
    Public Property Get NSU_MovAutomatizacao() As Long
        NSU_MovAutomatizacao = mNSU_MovAutomatizacao
    End Property
    
    Public Property Let IdEstabelecimento_MovAutomatizacao(ByVal Valor As Integer)
        mIdEstabelecimento_MovAutomatizacao = Valor
    End Property
    Public Property Get IdEstabelecimento_MovAutomatizacao() As Integer
        IdEstabelecimento_MovAutomatizacao = mIdEstabelecimento_MovAutomatizacao
    End Property
    
    Public Property Let DataSolicitacao_MovAutomatizacao(ByVal Valor As Date)
        mDataSolicitacao_MovAutomatizacao = Valor
    End Property
    Public Property Get DataSolicitacao_MovAutomatizacao() As Date
        DataSolicitacao_MovAutomatizacao = mDataSolicitacao_MovAutomatizacao
    End Property
    
    Public Property Let Operacao_MovAutomatizacao(ByVal Valor As String)
        mOperacao_MovAutomatizacao = Valor
    End Property
    Public Property Get Operacao_MovAutomatizacao() As String
        Operacao_MovAutomatizacao = mOperacao_MovAutomatizacao
    End Property
    
    Public Property Let Dados_MovAutomatizacao(ByVal Valor As String)
        mDados_MovAutomatizacao = Valor
    End Property
    Public Property Get Dados_MovAutomatizacao() As String
        Dados_MovAutomatizacao = mDados_MovAutomatizacao
    End Property
    
    Public Property Let IPComputadorAC_MovAutomatizacao(ByVal Valor As String)
        mIPComputadorAC_MovAutomatizacao = Valor
    End Property
    Public Property Get IPComputadorAC_MovAutomatizacao() As String
        IPComputadorAC_MovAutomatizacao = mIPComputadorAC_MovAutomatizacao
    End Property
    
    Public Property Let IPInternetAC_MovAutomatizacao(ByVal Valor As String)
        mIPInternetAC_MovAutomatizacao = Valor
    End Property
    Public Property Get IPInternetAC_MovAutomatizacao() As String
        IPInternetAC_MovAutomatizacao = mIPInternetAC_MovAutomatizacao
    End Property
    
    Public Property Let VersaoAC_MovAutomatizacao(ByVal Valor As String)
        mVersaoAC_MovAutomatizacao = Valor
    End Property
    Public Property Get VersaoAC_MovAutomatizacao() As String
        VersaoAC_MovAutomatizacao = mVersaoAC_MovAutomatizacao
    End Property
    
    Public Property Let VersaoHost_MovAutomatizacao(ByVal Valor As String)
        mVersaoHost_MovAutomatizacao = Valor
    End Property
    Public Property Get VersaoHost_MovAutomatizacao() As String
        VersaoHost_MovAutomatizacao = mVersaoHost_MovAutomatizacao
    End Property
    
    Public Property Let HoraAnalise_MovAutomatizacao(ByVal Valor As Date)
        mHoraAnalise_MovAutomatizacao = Valor
    End Property
    Public Property Get HoraAnalise_MovAutomatizacao() As Date
        HoraAnalise_MovAutomatizacao = mHoraAnalise_MovAutomatizacao
    End Property
    
    Public Property Let EtapaAndamento_MovAutomatizacao(ByVal Valor As Integer)
        mEtapaAndamento_MovAutomatizacao = Valor
    End Property
    Public Property Get EtapaAndamento_MovAutomatizacao() As Integer
        EtapaAndamento_MovAutomatizacao = mEtapaAndamento_MovAutomatizacao
    End Property
    
    Public Property Let DescricaoAndamento_MovAutomatizacao(ByVal Valor As String)
        mDescricaoAndamento_MovAutomatizacao = Valor
    End Property
    Public Property Get DescricaoAndamento_MovAutomatizacao() As String
        DescricaoAndamento_MovAutomatizacao = mDescricaoAndamento_MovAutomatizacao
    End Property
    
    Public Property Let HoraImportacaoDados_MovAutomatizacao(ByVal Valor As Date)
        mHoraImportacaoDados_MovAutomatizacao = Valor
    End Property
    Public Property Get HoraImportacaoDados_MovAutomatizacao() As Date
        HoraImportacaoDados_MovAutomatizacao = mHoraImportacaoDados_MovAutomatizacao
    End Property

    Public Property Let HoraImportacaoComplementar_MovAutomatizacao(ByVal Valor As Date)
        mHoraImportacaoComplementar_MovAutomatizacao = Valor
    End Property
    Public Property Get HoraImportacaoComplementar_MovAutomatizacao() As Date
        mHoraImportacaoComplementar_MovAutomatizacao = mHoraImportacaoComplementar_MovAutomatizacao
    End Property

    Public Property Let HoraConciliacaoDados_MovAutomatizacao(ByVal Valor As Date)
        mHoraConciliacaoDados_MovAutomatizacao = Valor
    End Property
    Public Property Get HoraConciliacaoDados_MovAutomatizacao() As Date
        HoraConciliacaoDados_MovAutomatizacao = mHoraConciliacaoDados_MovAutomatizacao
    End Property

    Public Property Let HoraCancelamentoHost_MovAutomatizacao(ByVal Valor As Date)
        mHoraCancelamentoHost_MovAutomatizacao = Valor
    End Property
    Public Property Get HoraCancelamentoHost_MovAutomatizacao() As Date
        HoraCancelamentoHost_MovAutomatizacao = mHoraCancelamentoHost_MovAutomatizacao
    End Property

    Public Property Let HoraConclusao_MovAutomatizacao(ByVal Valor As Date)
        mHoraConclusao_MovAutomatizacao = Valor
    End Property
    Public Property Get HoraConclusao_MovAutomatizacao() As Date
        HoraConclusao_MovAutomatizacao = mHoraConclusao_MovAutomatizacao
    End Property

    Public Property Let MensagemRetorno_MovAutomatizacao(ByVal Valor As String)
        mMensagemRetorno_MovAutomatizacao = Valor
    End Property
    Public Property Get MensagemRetorno_MovAutomatizacao() As String
        MensagemRetorno_MovAutomatizacao = mMensagemRetorno_MovAutomatizacao
    End Property

    Public Property Let Observacao_MovAutomatizacao(ByVal Valor As String)
        mObservacao_MovAutomatizacao = Valor
    End Property
    Public Property Get Observacao_MovAutomatizacao() As String
        Observacao_MovAutomatizacao = mObservacao_MovAutomatizacao
    End Property
    
Public Function Incluir() As Boolean

Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

        Incluir = False
        mNSU_MovAutomatizacao = ProximoRegistro()
        
        gSQL = "INSERT INTO MovAutomatizacao ( [NSU_MovAutomatizacao],[IdEstabelecimento_MovAutomatizacao] ,[DataSolicitacao_MovAutomatizacao] ,[Operacao_MovAutomatizacao] ,"
        gSQL = gSQL & "[Dados_MovAutomatizacao] ,[IPComputadorAC_MovAutomatizacao] ,[IPInternetAC_MovAutomatizacao] ,[VersaoAC_MovAutomatizacao] ,"
        gSQL = gSQL & "[VersaoHost_MovAutomatizacao] ,[HoraAnalise_MovAutomatizacao],[EtapaAndamento_MovAutomatizacao] ,[DescricaoAndamento_MovAutomatizacao] ,"
        gSQL = gSQL & "[HoraImportacaoDados_MovAutomatizacao],[HoraImportacaoComplementar_MovAutomatizacao] ,[HoraConciliacaoDados_MovAutomatizacao] ,"
        gSQL = gSQL & "[HoraCancelamentoHost_MovAutomatizacao] ,[HoraConclusao_MovAutomatizacao] ,[MensagemRetorno_MovAutomatizacao] ,"
        gSQL = gSQL & "[Observacao_MovAutomatizacao] ) VALUES ( "
        Call sqlNumero(1, CStr(mNSU_MovAutomatizacao), ", ")
        Call sqlNumero(1, CStr(mIdEstabelecimento_MovAutomatizacao), ", ")
        Call sqlDataHora(1, CStr(mDataSolicitacao_MovAutomatizacao), ", ")
        Call sqlTexto(1, mOperacao_MovAutomatizacao, ", ")
        Call sqlTexto(1, CStr(mDados_MovAutomatizacao), ", ")
        Call sqlTexto(1, mIPComputadorAC_MovAutomatizacao, ", ")
        Call sqlTexto(1, mIPInternetAC_MovAutomatizacao, ", ")
        Call sqlTexto(1, mVersaoAC_MovAutomatizacao, ", ")
        Call sqlTexto(1, mVersaoHost_MovAutomatizacao, ", ")
        Call sqlHora(1, mHoraAnalise_MovAutomatizacao, ", ")
        Call sqlNumero(1, mEtapaAndamento_MovAutomatizacao, ", ")
        Call sqlTexto(1, CStr(mDescricaoAndamento_MovAutomatizacao), ", ")
        Call sqlHora(1, mHoraImportacaoDados_MovAutomatizacao, ", ")
        Call sqlHora(1, mHoraImportacaoComplementar_MovAutomatizacao, ", ")
        Call sqlHora(1, mHoraConciliacaoDados_MovAutomatizacao, ", ")
        Call sqlHora(1, CStr(mHoraCancelamentoHost_MovAutomatizacao), ", ")
        Call sqlHora(1, CStr(mHoraConclusao_MovAutomatizacao), ", ")
        Call sqlTexto(1, CStr(mMensagemRetorno_MovAutomatizacao), ", ")
        Call sqlTexto(1, CStr(mObservacao_MovAutomatizacao), " )")
        
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
            
        If lRecordsAffected > 0 Then
           Incluir = True
        End If
        Exit Function
Err_Incluir:
    Call CriaLogCadastroDll2("cMovAutomatizacao:Incluir - Erro desconhecido.", Err.Description, gSQL)
End Function
    
Public Function ProximoRegistro() As Long
    
    
     On Error GoTo trata_erro
            
            ProximoRegistro = 1
            gSQL = ""
            
            gSQL = "SELECT NSU_MovAutomatizacao"
            gSQL = gSQL & "  FROM MovAutomatizacao"
            gSQL = gSQL & " ORDER BY NSU_MovAutomatizacao DESC"
            If Localizar(1) Then
                ProximoRegistro = rsAutomatizacao("NSU_MovAutomatizacao") + 1
            End If
            rsAutomatizacao.Close
            Set rsAutomatizacao = Nothing
            Exit Function
        
trata_erro:
            MsgBox Err.Number & " - " & Err.Description
    
End Function
    
Public Function DefineHoraAnalise(ByVal pNSU As Long, ByVal pHoraAnalise As Date, ByVal pVersaoHost As String) As Boolean
        
        
Dim lRecordsAffected As Long
    
On Error GoTo Err_DefineHoraAnalise
    
        DefineHoraAnalise = False
        gSQL = "UPDATE MovAutomatizacao SET "
        Call sqlHora(2, "HoraAnalise_MovAutomatizacao = ", pHoraAnalise)
        Call sqlTexto(2, ", VersaoHost_MovAutomatizacao = ", pVersaoHost)
        gSQL = gSQL & " WHERE NSU_MovAutomatizacao = " & pNSU
        gSQL = gSQL & " AND HoraAnalise_MovAutomatizacao IS NULL"
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        
        If lRecordsAffected > 0 Then
            DefineHoraAnalise = True
        End If
        Exit Function
    
Err_DefineHoraAnalise:

End Function
Public Function DefineHoraConclusao(ByVal pNSU As Long, ByVal pHoraConclusao As Date, ByVal pVersaoHost As String) As Boolean
        
Dim lRecordsAffected As Long
    
On Error GoTo Err_DefineHoraConclusao
    
        DefineHoraConclusao = False
        gSQL = "UPDATE MovAutomatizacao SET "
        Call sqlHora(2, "HoraAprovacao_MovAutomatizacao = ", pHoraConclusao)
        Call sqlTexto(2, ", VersaoHost_MovAutomatizacao = ", pVersaoHost)
        gSQL = gSQL & " WHERE NSU_MovAutomatizacao = " & pNSU
        gSQL = gSQL & " AND HoraConclusao_MovAutomatizacao IS NULL"
        
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        
        If lRecordsAffected > 0 Then
            DefineHoraConclusao = True
        End If
        Exit Function
    
Err_DefineHoraConclusao:

End Function
Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsAutomatizacao = New adodb.Recordset
    rsAutomatizacao.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsAutomatizacao.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsAutomatizacao.EOF Then
        AtualizaRecordset = True
    End If
End Function
    
   Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
        Localizar = False
        If AtualizaRecordset(pQtdRegistro) Then
            Localizar = True
        End If
    End Function
    
    Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
        PreparaSQL = Trim("SELECT [NSU_MovAutomatizacao],[IdEstabelecimento_MovAutomatizacao] ,[DataSolicitacao_MovAutomatizacao] ,[Operacao_MovAutomatizacao] , [Dados_MovAutomatizacao] ,[IPComputadorAC_MovAutomatizacao] ,[IPInternetAC_MovAutomatizacao] ,[VersaoAC_MovAutomatizacao] ,[VersaoHost_MovAutomatizacao] ,[HoraAnalise_MovAutomatizacao],[EtapaAndamento_MovAutomatizacao] ,[DescricaoAndamento_MovAutomatizacao] ,[HoraImportacaoDados_MovAutomatizacao],[HoraImportacaoComplementar_MovAutomatizacao] ,[HoraConciliacaoDados_MovAutomatizacao] ,[HoraCancelamentoHost_MovAutomatizacao] ,[HoraConclusao_MovAutomatizacao] ,[MensagemRetorno_MovAutomatizacao] ,[Observacao_MovAutomatizacao] FROM MovAutomatizacao" & " " & xCondicao & " " & xTipoOrdem)
    End Function

    
