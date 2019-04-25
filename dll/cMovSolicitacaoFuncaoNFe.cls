VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovSolicitacaoFuncaoNFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Private mNSU_MovSolicitacaoFuncaoNFe As Long
    Private mNumeroControleSolicitacao_MovSolicitacaoFuncaoNFe As Long
    Private mDataSolicitacao_MovSolicitacaoFuncaoNFe As Date
    Private mTipoOperacao_MovSolicitacaoFuncaoNFe As String
    Private mCodigoEstabelecimento_MovSolicitacaoFuncaoNFe As Integer
    Private mSerieNFe_MovSolicitacaoFuncaoNFe As String
    Private mNumeroNFe_MovSolicitacaoFuncaoNFe As Long
    Private mChaveAcessoNFe_MovSolicitacaoFuncaoNFe As String
    Private mIPComputadorAC_MovSolicitacaoFuncaoNFe As String
    Private mIPInternetAC_MovSolicitacaoFuncaoNFe As String
    Private mSegurancaEstabelecimento_MovSolicitacaoFuncaoNFe As String
    Private mCodigoUsuario_MovSolicitacaoFuncaoNFe As Integer
    Private mVersaoAC_MovSolicitacaoFuncaoNFe As String
    Private mVersaoHost_MovSolicitacaoFuncaoNFe As String
    Private mTexto_MovSolicitacaoFuncaoNFe As String
    Private mHoraAnalise_MovSolicitacaoFuncaoNFe As Date
    Private mHoraAprovacao_MovSolicitacaoFuncaoNFe As Date
    Private mHoraCancelamentoHost_MovSolicitacaoFuncaoNFe As Date
    Private mHoraConfirmacaoAC_MovSolicitacaoFuncaoNFe As Date
    Private mHoraCancelamentoAC_MovSolicitacaoFuncaoNFe As Date
    Private mMensagem_MovSolicitacaoFuncaoNFe As String
    Private mCodigoRetorno_MovSolicitacaoFuncaoNFe As Integer
    Private mNumeroLote_MovSolicitacaoFuncaoNFe As String

    Private rsSolicitacaoFuncaoNFe As New adodb.Recordset

    Public Property Let NSU_MovSolicitacaoFuncaoNFe(ByVal Valor As Long)
        mNSU_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get NSU_MovSolicitacaoFuncaoNFe() As Long
        NSU_MovSolicitacaoFuncaoNFe = mNSU_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let NumeroControleSolicitacao_MovSolicitacaoFuncaoNFe(ByVal Valor As Long)
        mNumeroControleSolicitacao_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get NumeroControleSolicitacao_MovSolicitacaoFuncaoNFe() As Long
        NumeroControleSolicitacao_MovSolicitacaoFuncaoNFe = mNumeroControleSolicitacao_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let DataSolicitacao_MovSolicitacaoFuncaoNFe(ByVal Valor As Date)
        mDataSolicitacao_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get DataSolicitacao_MovSolicitacaoFuncaoNFe() As Date
        DataSolicitacao_MovSolicitacaoFuncaoNFe = mDataSolicitacao_MovSolicitacaoFuncaoNFe
    End Property
  
    Public Property Let TipoOperacao_MovSolicitacaoFuncaoNFe(ByVal Valor As String)
        mTipoOperacao_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get TipoOperacao_MovSolicitacaoFuncaoNFe() As String
        TipoOperacao_MovSolicitacaoFuncaoNFe = mTipoOperacao_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let CodigoEstabelecimento_MovSolicitacaoFuncaoNFe(ByVal Valor As Integer)
        mCodigoEstabelecimento_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get CodigoEstabelecimento_MovSolicitacaoFuncaoNFe() As Integer
        CodigoEstabelecimento_MovSolicitacaoFuncaoNFe = mCodigoEstabelecimento_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let SerieNFe_MovSolicitacaoFuncaoNFe(ByVal Valor As String)
        mSerieNFe_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get SerieNFe_MovSolicitacaoFuncaoNFe() As String
        SerieNFe_MovSolicitacaoFuncaoNFe = mSerieNFe_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let NumeroNFe_MovSolicitacaoFuncaoNFe(ByVal Valor As Long)
        mNumeroNFe_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get NumeroNFe_MovSolicitacaoFuncaoNFe() As Long
        NumeroNFe_MovSolicitacaoFuncaoNFe = mNumeroNFe_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let ChaveAcessoNFe_MovSolicitacaoFuncaoNFe(ByVal Valor As String)
        mChaveAcessoNFe_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get ChaveAcessoNFe_MovSolicitacaoFuncaoNFe() As String
        ChaveAcessoNFe_MovSolicitacaoFuncaoNFe = mChaveAcessoNFe_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let IPComputadorAC_MovSolicitacaoFuncaoNFe(ByVal Valor As String)
        mIPComputadorAC_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get IPComputadorAC_MovSolicitacaoFuncaoNFe() As String
        IPComputadorAC_MovSolicitacaoFuncaoNFe = mIPComputadorAC_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let IPInternetAC_MovSolicitacaoFuncaoNFe(ByVal Valor As String)
        mIPInternetAC_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get IPInternetAC_MovSolicitacaoFuncaoNFe() As String
        IPInternetAC_MovSolicitacaoFuncaoNFe = mIPInternetAC_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let SegurancaEstabelecimento_MovSolicitacaoFuncaoNFe(ByVal Valor As String)
        mSegurancaEstabelecimento_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get SegurancaEstabelecimento_MovSolicitacaoFuncaoNFe() As String
        SegurancaEstabelecimento_MovSolicitacaoFuncaoNFe = mSegurancaEstabelecimento_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let CodigoUsuario_MovSolicitacaoFuncaoNFe(ByVal Valor As Integer)
        mCodigoUsuario_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get CodigoUsuario_MovSolicitacaoFuncaoNFe() As Integer
        CodigoUsuario_MovSolicitacaoFuncaoNFe = mCodigoUsuario_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let VersaoAC_MovSolicitacaoFuncaoNFe(ByVal Valor As String)
        mVersaoAC_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get VersaoAC_MovSolicitacaoFuncaoNFe() As String
        VersaoAC_MovSolicitacaoFuncaoNFe = mVersaoAC_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let VersaoHost_MovSolicitacaoFuncaoNFe(ByVal Valor As String)
        mVersaoHost_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get VersaoHost_MovSolicitacaoFuncaoNFe() As String
        VersaoHost_MovSolicitacaoFuncaoNFe = mVersaoHost_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let Texto_MovSolicitacaoFuncaoNFe(ByVal Valor As String)
        mTexto_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get Texto_MovSolicitacaoFuncaoNFe() As String
        Texto_MovSolicitacaoFuncaoNFe = mTexto_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let HoraAnalise_MovSolicitacaoFuncaoNFe(ByVal Valor As Date)
        mHoraAnalise_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get HoraAnalise_MovSolicitacaoFuncaoNFe() As Date
        HoraAnalise_MovSolicitacaoFuncaoNFe = mHoraAnalise_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let HoraAprovacao_MovSolicitacaoFuncaoNFe(ByVal Valor As Date)
        mHoraAprovacao_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get HoraAprovacao_MovSolicitacaoFuncaoNFe() As Date
        HoraAprovacao_MovSolicitacaoFuncaoNFe = mHoraAprovacao_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let HoraCancelamentoHost_MovSolicitacaoFuncaoNFe(ByVal Valor As Date)
        mHoraCancelamentoHost_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get HoraCancelamentoHost_MovSolicitacaoFuncaoNFe() As Date
        HoraCancelamentoHost_MovSolicitacaoFuncaoNFe = mHoraCancelamentoHost_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe(ByVal Valor As Date)
        mHoraConfirmacaoAC_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe() As Date
        HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe = mHoraConfirmacaoAC_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let HoraCancelamentoAC_MovSolicitacaoFuncaoNFe(ByVal Valor As Date)
        mHoraCancelamentoAC_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get HoraCancelamentoAC_MovSolicitacaoFuncaoNFe() As Date
        HoraCancelamentoAC_MovSolicitacaoFuncaoNFe = mHoraCancelamentoAC_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let Mensagem_MovSolicitacaoFuncaoNFe(ByVal Valor As String)
        mMensagem_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get Mensagem_MovSolicitacaoFuncaoNFe() As String
        Mensagem_MovSolicitacaoFuncaoNFe = mMensagem_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let CodigoRetorno_MovSolicitacaoFuncaoNFe(ByVal Valor As Integer)
        mCodigoRetorno_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get CodigoRetorno_MovSolicitacaoFuncaoNFe() As Integer
        CodigoRetorno_MovSolicitacaoFuncaoNFe = mCodigoRetorno_MovSolicitacaoFuncaoNFe
    End Property
    
    Public Property Let NumeroLote_MovSolicitacaoFuncaoNFe(ByVal Valor As String)
        mNumeroLote_MovSolicitacaoFuncaoNFe = Valor
    End Property
    Public Property Get NumeroLote_MovSolicitacaoFuncaoNFe() As String
        NumeroLote_MovSolicitacaoFuncaoNFe = mNumeroLote_MovSolicitacaoFuncaoNFe
    End Property
    
    
    Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
        Dim i As Integer
        AtualizaRecordset = False
        Set rsSolicitacaoFuncaoNFe = New adodb.Recordset
        rsSolicitacaoFuncaoNFe.CursorLocation = adUseClient
        i = Len(gSQL)
        If pQtdRegistro > 0 Then
            gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
        End If
        rsSolicitacaoFuncaoNFe.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
        If Not rsSolicitacaoFuncaoNFe.EOF Then
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
        PreparaSQL = Trim("SELECT NSU_MovSolicitacaoFuncaoNFe, NumeroControleSolicitacao_MovSolicitacaoFuncaoNFe, DataSolicitacao_MovSolicitacaoFuncaoNFe, TipoOperacao_MovSolicitacaoFuncaoNFe, CodigoEstabelecimento_MovSolicitacaoFuncaoNFe, SerieNFe_MovSolicitacaoFuncaoNFe, NumeroNFe_MovSolicitacaoFuncaoNFe, ChaveAcessoNFe_MovSolicitacaoFuncaoNFe, IPComputadorAC_MovSolicitacaoFuncaoNFe, IPInternetAC_MovSolicitacaoFuncaoNFe, SegurancaEstabelecimento_MovSolicitacaoFuncaoNFe, CodigoUsuario_MovSolicitacaoFuncaoNFe, VersaoAC_MovSolicitacaoFuncaoNFe, VersaoHost_MovSolicitacaoFuncaoNFe, Texto_MovSolicitacaoFuncaoNFe, HoraAnalise_MovSolicitacaoFuncaoNFe, HoraAprovacao_MovSolicitacaoFuncaoNFe, HoraCancelamentoHost_MovSolicitacaoFuncaoNFe, HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe , HoraCancelamentoAC_MovSolicitacaoFuncaoNFe, Mensagem_MovSolicitacaoFuncaoNFe, CodigoRetorno_MovSolicitacaoFuncaoNFe, NumeroLote_MovSolicitacaoFuncaoNFe FROM MovSolicitacaoFuncaoNFe" & " " & xCondicao & " " & xTipoOrdem)
    End Function

Public Function BuscaMensagem(ByVal pNSU As Long) As String
  
  On Error GoTo trata_erro
        
        BuscaMensagem = ""
        gSQL = ""
        
        gSQL = "SELECT Mensagem_MovSolicitacaoFuncaoNFe"
        gSQL = gSQL & "  FROM MovSolicitacaoFuncaoNFe"
        gSQL = gSQL & " WHERE NSU_MovSolicitacaoFuncaoNFe = " & pNSU
        
        If Localizar(1) Then
            BuscaMensagem = rsSolicitacaoFuncaoNFe("Mensagem_MovSolicitacaoFuncaoNFe")
        End If
        rsSolicitacaoFuncaoNFe.Close
        Set rsSolicitacaoFuncaoNFe = Nothing
        Exit Function
        
trata_erro:
        MsgBox Err.Number & " - " & Err.Description
    
End Function
Public Function DefineHoraAnalise(ByVal pNSU As Long, ByVal pHoraAnalise As Date, ByVal pVersaoHost As String) As Boolean
        
        
Dim lRecordsAffected As Long
    
On Error GoTo Err_DefineHoraAnalise
    
        DefineHoraAnalise = False
        gSQL = "UPDATE MovSolicitacaoFuncaoNFe SET "
        Call sqlHora(2, "HoraAnalise_MovSolicitacaoFuncaoNFe = ", pHoraAnalise)
        Call sqlTexto(2, ", VersaoHost_MovSolicitacaoFuncaoNFe = ", pVersaoHost)
        gSQL = gSQL & " WHERE NSU_MovSolicitacaoFuncaoNFe = " & pNSU
        gSQL = gSQL & " AND HoraAnalise_MovSolicitacaoFuncaoNFe IS NULL"
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        
        If lRecordsAffected > 0 Then
            DefineHoraAnalise = True
        End If
        Exit Function
    
Err_DefineHoraAnalise:

End Function
Public Function DefineHoraAprovacao(ByVal pNSU As Long, ByVal pHoraAprovacao As Date, ByVal pVersaoHost As String) As Boolean
        
Dim lRecordsAffected As Long
    
On Error GoTo Err_DefineHoraAprovacao
    
        DefineHoraAprovacao = False
        gSQL = "UPDATE MovSolicitacaoFuncaoNFe SET "
        Call sqlHora(2, "HoraAprovacao_MovSolicitacaoFuncaoNFe = ", pHoraAprovacao)
        Call sqlTexto(2, ", VersaoHost_MovSolicitacaoFuncaoNFe = ", pVersaoHost)
        gSQL = gSQL & " WHERE NSU_MovSolicitacaoFuncaoNFe = " & pNSU
        gSQL = gSQL & " AND HoraAprovacao_MovSolicitacaoFuncaoNFe IS NULL"
        
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        
        If lRecordsAffected > 0 Then
            DefineHoraAprovacao = True
        End If
        Exit Function
    
Err_DefineHoraAprovacao:

End Function
    
Public Function DefineHoraCancelamentoAC(ByVal pNSU As Long, ByVal pHoraCancelamentoAC As Date) As Boolean
        
        
Dim lRecordsAffected As Long
    
On Error GoTo Err_DefineHoraCancelamentoAC
    
        DefineHoraCancelamentoAC = False
        gSQL = "UPDATE MovSolicitacaoFuncaoNFe SET "
        Call sqlHora(2, "HoraCancelamentoAC_MovSolicitacaoFuncaoNFe = ", pHoraCancelamentoAC)
        gSQL = gSQL & " WHERE NSU_MovSolicitacaoFuncaoNFe = " & pNSU
        
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        
        If lRecordsAffected > 0 Then
            DefineHoraCancelamentoAC = True
        End If
        Exit Function
    
Err_DefineHoraCancelamentoAC:
        
End Function
Public Function DefineHoraCancelamentoHost(ByVal pNSU As Long, ByVal pHoraCancelamentoHost As Date, ByVal pVersaoHost As String) As Boolean
        
        Dim lRecordsAffected As Long
    
On Error GoTo Err_DefineHoraCancelamentoHost
    
        DefineHoraCancelamentoHost = False
        gSQL = "UPDATE MovSolicitacaoFuncaoNFe SET "
        Call sqlHora(2, "HoraCancelamentoHost_MovSolicitacaoFuncaoNFe = ", pHoraCancelamentoHost)
        Call sqlTexto(2, ", VersaoHost_MovSolicitacaoFuncaoNFe = ", pVersaoHost)
        gSQL = gSQL & " WHERE NSU_MovSolicitacaoFuncaoNFe = " & pNSU
        gSQL = gSQL & " AND HoraCancelamentoHost_MovSolicitacaoFuncaoNFe IS NULL"
        
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        
        If lRecordsAffected > 0 Then
            DefineHoraCancelamentoHost = True
        End If
        Exit Function
    
Err_DefineHoraCancelamentoHost:

End Function
Public Function DefineHoraConfirmacaoAC(ByVal pNSU As Long, ByVal pHoraConfirmacaoAC As Date) As Boolean
        
        
Dim lRecordsAffected As Long
    
On Error GoTo Err_DefineHoraConfirmacaoAC
    
        DefineHoraConfirmacaoAC = False
        gSQL = "UPDATE MovSolicitacaoFuncaoNFe SET "
        Call sqlHora(2, "HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe = ", pHoraConfirmacaoAC)
        gSQL = gSQL & " WHERE NSU_MovSolicitacaoFuncaoNFe = " & pNSU
        gSQL = gSQL & " AND HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe IS NULL"

        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        
        If lRecordsAffected > 0 Then
            DefineHoraConfirmacaoAC = True
        End If
        Exit Function
    
Err_DefineHoraConfirmacaoAC:

End Function
Public Function DefineMensagem(ByVal pNSU As Long, ByVal pMensagem As String) As Boolean
        
Dim lRecordsAffected As Long
    
On Error GoTo Err_DefineMensagem
    
        DefineMensagem = False
        gSQL = "UPDATE MovSolicitacaoFuncaoNFe SET "
        Call sqlHora(2, "Mensagem_MovSolicitacaoFuncaoNFe = ", pMensagem)
        gSQL = gSQL & " WHERE NSU_MovSolicitacaoFuncaoNFe = " & pNSU

        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        
        If lRecordsAffected > 0 Then
            DefineMensagem = True
        End If
        Exit Function
    
Err_DefineMensagem:

        
End Function
Public Function Alterar(ByVal pNSU As Long) As Boolean

        Dim lRecordsAffected As Long
        
        On Error GoTo Err_Alterar
            Alterar = False
            gSQL = "UPDATE MovSolicitacaoFuncaoNFe SET "
            Call sqlNumero(2, "NSU_MovSolicitacaoFuncaoNFe = ", CStr(mNSU_MovSolicitacaoFuncaoNFe))
            Call sqlNumero(2, ", NumeroControleSolicitacao_MovSolicitacaoFuncaoNFe = ", CStr(mNumeroControleSolicitacao_MovSolicitacaoFuncaoNFe))
            Call sqlDataHora(2, ", DataSolicitacao_MovSolicitacaoFuncaoNFe = ", CStr(mDataSolicitacao_MovSolicitacaoFuncaoNFe))
            Call sqlTexto(2, ", TipoOperacao_MovSolicitacaoFuncaoNFe = ", mTipoOperacao_MovSolicitacaoFuncaoNFe)
            Call sqlNumero(2, ", CodigoEstabelecimento_MovSolicitacaoFuncaoNFe = ", CStr(mCodigoEstabelecimento_MovSolicitacaoFuncaoNFe))
            Call sqlTexto(2, ", SerieNFe_MovSolicitacaoFuncaoNFe = ", mSerieNFe_MovSolicitacaoFuncaoNFe)
            Call sqlNumero(2, ", NumeroNFe_MovSolicitacaoFuncaoNFe = ", CStr(mNumeroNFe_MovSolicitacaoFuncaoNFe))
            Call sqlTexto(2, ", ChaveAcessoNFe_MovSolicitacaoFuncaoNFe = ", mChaveAcessoNFe_MovSolicitacaoFuncaoNFe)
            Call sqlTexto(2, ", IPComputadorAC_MovSolicitacaoFuncaoNFe = ", mIPComputadorAC_MovSolicitacaoFuncaoNFe)
            Call sqlTexto(2, ", IPInternetAC_MovSolicitacaoFuncaoNFe = ", mIPInternetAC_MovSolicitacaoFuncaoNFe)
            Call sqlTexto(2, ", SegurancaEstabelecimento_MovSolicitacaoFuncaoNFe = ", mSegurancaEstabelecimento_MovSolicitacaoFuncaoNFe)
            Call sqlNumero(2, ", CodigoUsuario_MovSolicitacaoFuncaoNFe = ", CStr(mCodigoUsuario_MovSolicitacaoFuncaoNFe))
            Call sqlTexto(2, ", VersaoAC_MovSolicitacaoFuncaoNFe = ", mVersaoAC_MovSolicitacaoFuncaoNFe)
            Call sqlTexto(2, ", VersaoHost_MovSolicitacaoFuncaoNFe = ", mVersaoHost_MovSolicitacaoFuncaoNFe)
            Call sqlTexto(2, ", Texto_MovSolicitacaoFuncaoNFe = ", mTexto_MovSolicitacaoFuncaoNFe)
            Call sqlHora(2, ", HoraAnalise_MovSolicitacaoFuncaoNFe = ", CStr(mHoraAnalise_MovSolicitacaoFuncaoNFe))
            Call sqlHora(2, ", HoraAprovacao_MovSolicitacaoFuncaoNFe = ", CStr(mHoraAprovacao_MovSolicitacaoFuncaoNFe))
            Call sqlHora(2, ", HoraCancelamentoHost_MovSolicitacaoFuncaoNFe = ", CStr(mHoraCancelamentoHost_MovSolicitacaoFuncaoNFe))
            Call sqlHora(2, ", HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe = ", CStr(mHoraConfirmacaoAC_MovSolicitacaoFuncaoNFe))
            Call sqlHora(2, ", HoraCancelamentoAC_MovSolicitacaoFuncaoNFe = ", CStr(mHoraCancelamentoAC_MovSolicitacaoFuncaoNFe))
            Call sqlTexto(2, ", Mensagem_MovSolicitacaoFuncaoNFe = ", mMensagem_MovSolicitacaoFuncaoNFe)
            Call sqlNumero(2, ", CodigoRetorno_MovSolicitacaoFuncaoNFe = ", CStr(mCodigoRetorno_MovSolicitacaoFuncaoNFe))
            Call sqlTexto(2, ", NumeroLote_MovSolicitacaoFuncaoNFe = ", mNumeroLote_MovSolicitacaoFuncaoNFe)
            gSQL = gSQL & " WHERE NSU_MovSolicitacaoFuncaoNFe = " & pNSU
            
            gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
            
            If lRecordsAffected > 0 Then
                Alterar = True
            End If
            
Err_Alterar:
            
End Function


Public Function LocalizarNSU(ByVal pEmpresa As Integer, ByVal pNSU As Long) As String
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarNSU = False
    xCondicao = " WHERE CodigoEstabelecimento_MovSolicitacaoFuncaoNFe = " & pEmpresa
    xCondicao = xCondicao & " AND NSU_MovSolicitacaoFuncaoNFe = " & pNSU
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarNSU = Localizar(1)
    If LocalizarNSU Then
        AtribuiValor
    End If
    rsSolicitacaoFuncaoNFe.Close
    Set rsSolicitacaoFuncaoNFe = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Excluir(ByVal pNSU As Long) As Boolean
    
    
    Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM MovSolicitacaoFuncaoNFe"
    gSQL = gSQL & " WHERE NSU_MovSolicitacaoFuncaoNFe = " & pNSU
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
    
End Function
Public Function Incluir() As Boolean

Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

        Incluir = False
        mNSU_MovSolicitacaoFuncaoNFe = ProximoRegistroNFCE()
        
        gSQL = "INSERT INTO MovSolicitacaoFuncaoNFe ( NSU_MovSolicitacaoFuncaoNFe, NumeroControleSolicitacao_MovSolicitacaoFuncaoNFe, "
        gSQL = gSQL & "DataSolicitacao_MovSolicitacaoFuncaoNFe, TipoOperacao_MovSolicitacaoFuncaoNFe, CodigoEstabelecimento_MovSolicitacaoFuncaoNFe, "
        gSQL = gSQL & "SerieNFe_MovSolicitacaoFuncaoNFe, NumeroNFe_MovSolicitacaoFuncaoNFe, ChaveAcessoNFe_MovSolicitacaoFuncaoNFe, "
        gSQL = gSQL & "IPComputadorAC_MovSolicitacaoFuncaoNFe, IPInternetAC_MovSolicitacaoFuncaoNFe, SegurancaEstabelecimento_MovSolicitacaoFuncaoNFe, "
        gSQL = gSQL & "CodigoUsuario_MovSolicitacaoFuncaoNFe, VersaoAC_MovSolicitacaoFuncaoNFe, VersaoHost_MovSolicitacaoFuncaoNFe, "
        gSQL = gSQL & "Texto_MovSolicitacaoFuncaoNFe, HoraAnalise_MovSolicitacaoFuncaoNFe, HoraAprovacao_MovSolicitacaoFuncaoNFe, "
        gSQL = gSQL & "HoraCancelamentoHost_MovSolicitacaoFuncaoNFe, HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe , "
        gSQL = gSQL & "HoraCancelamentoAC_MovSolicitacaoFuncaoNFe, Mensagem_MovSolicitacaoFuncaoNFe, "
        gSQL = gSQL & "CodigoRetorno_MovSolicitacaoFuncaoNFe, NumeroLote_MovSolicitacaoFuncaoNFe ) VALUES ( "
        Call sqlNumero(1, CStr(mNSU_MovSolicitacaoFuncaoNFe), ", ")
        Call sqlNumero(1, CStr(mNumeroControleSolicitacao_MovSolicitacaoFuncaoNFe), ", ")
        Call sqlDataHora(1, CStr(mDataSolicitacao_MovSolicitacaoFuncaoNFe), ", ")
        Call sqlTexto(1, mTipoOperacao_MovSolicitacaoFuncaoNFe, ", ")
        Call sqlNumero(1, CStr(mCodigoEstabelecimento_MovSolicitacaoFuncaoNFe), ", ")
        Call sqlTexto(1, mSerieNFe_MovSolicitacaoFuncaoNFe, ", ")
        Call sqlTexto(1, mNumeroNFe_MovSolicitacaoFuncaoNFe, ", ")
        Call sqlTexto(1, mChaveAcessoNFe_MovSolicitacaoFuncaoNFe, ", ")
        Call sqlTexto(1, mIPComputadorAC_MovSolicitacaoFuncaoNFe, ", ")
        Call sqlTexto(1, mIPInternetAC_MovSolicitacaoFuncaoNFe, ", ")
        Call sqlTexto(1, mSegurancaEstabelecimento_MovSolicitacaoFuncaoNFe, ", ")
        Call sqlNumero(1, CStr(mSegurancaEstabelecimento_MovSolicitacaoFuncaoNFe), ", ")
        Call sqlTexto(1, mVersaoAC_MovSolicitacaoFuncaoNFe, ", ")
        Call sqlTexto(1, mVersaoHost_MovSolicitacaoFuncaoNFe, ", ")
        Call sqlTexto(1, mTexto_MovSolicitacaoFuncaoNFe, ", ")
        Call sqlHora(1, CStr(mHoraAnalise_MovSolicitacaoFuncaoNFe), ", ")
        Call sqlHora(1, CStr(mHoraAprovacao_MovSolicitacaoFuncaoNFe), ", ")
        Call sqlHora(1, CStr(mHoraCancelamentoHost_MovSolicitacaoFuncaoNFe), ", ")
        Call sqlHora(1, CStr(mHoraConfirmacaoAC_MovSolicitacaoFuncaoNFe), ", ")
        Call sqlHora(1, CStr(mHoraCancelamentoAC_MovSolicitacaoFuncaoNFe), ", ")
        Call sqlTexto(1, mMensagem_MovSolicitacaoFuncaoNFe, ", ")
        Call sqlNumero(1, CStr(mCodigoRetorno_MovSolicitacaoFuncaoNFe), ", ")
        Call sqlTexto(1, mNumeroLote_MovSolicitacaoFuncaoNFe, " )")
        
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
            
        If lRecordsAffected > 0 Then
           Incluir = True
        End If
        Exit Function
Err_Incluir:
    Call CriaLogCadastroDll2("cMovSolicitacaoFuncaoNFe:Incluir - Erro desconhecido.", Err.Description, gSQL)
End Function
Public Function ProximoRegistroNFCE() As Long
    
    
     On Error GoTo trata_erro
            
            ProximoRegistroNFCE = 1
            gSQL = ""
            
            gSQL = "SELECT NSU_MovSolicitacaoFuncaoNFe"
            gSQL = gSQL & "  FROM MovSolicitacaoFuncaoNFe"
            gSQL = gSQL & " ORDER BY NSU_MovSolicitacaoFuncaoNFe DESC"
            If Localizar(1) Then
                ProximoRegistroNFCE = rsSolicitacaoFuncaoNFe("NSU_MovSolicitacaoFuncaoNFe") + 1
            End If
            rsSolicitacaoFuncaoNFe.Close
            Set rsSolicitacaoFuncaoNFe = Nothing
            Exit Function
        
trata_erro:
            MsgBox Err.Number & " - " & Err.Description
    
End Function
Public Function ProximoNumeroNFCE() As Long
    
    
     On Error GoTo trata_erro
            
            ProximoNumeroNFCE = 1
            gSQL = ""
            
            gSQL = "SELECT [NumeroNFe_MovSolicitacaoFuncaoNFe]"
            gSQL = gSQL & "  FROM MovSolicitacaoFuncaoNFe"
            gSQL = gSQL & "  WHERE TipoOperacao_MovSolicitacaoFuncaoNFe = " & preparaTexto("NFCe 3.10")
            gSQL = gSQL & " ORDER BY NSU_MovSolicitacaoFuncaoNFe DESC"
            If Localizar(1) Then
                ProximoNumeroNFCE = rsSolicitacaoFuncaoNFe("NumeroNFe_MovSolicitacaoFuncaoNFe") + 1
            End If
            rsSolicitacaoFuncaoNFe.Close
            Set rsSolicitacaoFuncaoNFe = Nothing
            Exit Function
        
trata_erro:
            MsgBox Err.Number & " - " & Err.Description
    
End Function
Public Function ExisteSolicitacaoParaOperacao(ByVal pData As Date, ByVal pNumeroNFe As Long, ByVal pSerie As Integer, ByVal pTipoOperacao As String) As Boolean
     On Error GoTo trata_erro
        ExisteSolicitacaoParaOperacao = False
        gSQL = ""
        
        gSQL = "SELECT NSU_MovSolicitacaoFuncaoNFe"
        gSQL = gSQL & "  FROM MovSolicitacaoFuncaoNFe"
        gSQL = gSQL & " WHERE TipoOperacao_MovSolicitacaoFuncaoNFe = " & preparaTexto(pTipoOperacao)
        gSQL = gSQL & "   AND DataSolicitacao_MovSolicitacaoFuncaoNFe = " & preparaData(pData)
        gSQL = gSQL & "   AND NumeroNFe_MovSolicitacaoFuncaoNFe = " & pNumeroNFe
        gSQL = gSQL & "   AND SerieNFe_MovSolicitacaoFuncaoNFe = " & pSerie
        
        ExisteSolicitacaoParaOperacao = Localizar(1)
        rsSolicitacaoFuncaoNFe.Close
        Set rsSolicitacaoFuncaoNFe = Nothing
        
        Exit Function
    


trata_erro:
    MsgBox Err.Number & " - " & Err.Description

End Function


Public Function VerificaSeEstaEmAnalise(ByVal pNSU As Long) As Boolean

On Error GoTo trata_erro
    
    VerificaSeEstaEmAnalise = False
    gSQL = "SELECT NSU_MovSolicitacaoFuncaoNFe"
    gSQL = gSQL & "  FROM MovSolicitacaoFuncaoNFe"
    gSQL = gSQL & " WHERE NSU_MovSolicitacaoFuncaoNFe = " & pNSU
    gSQL = gSQL & "   AND NOT HoraAnalise_MovSolicitacaoFuncaoNFe IS NULL"
    VerificaSeEstaEmAnalise = Localizar(1)
    Exit Function

trata_erro:
End Function

Private Sub AtribuiValor()

    Dim xLocal As Integer
    
    On Error GoTo trata_erro
            
    xLocal = 1
    NSU_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("NSU_MovSolicitacaoFuncaoNFe").Value
    xLocal = 2
    NumeroControleSolicitacao_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("NumeroControleSolicitacao_MovSolicitacaoFuncaoNFe").Value
    xLocal = 3
    DataSolicitacao_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("DataSolicitacao_MovSolicitacaoFuncaoNFe").Value
    xLocal = 4
    TipoOperacao_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("TipoOperacao_MovSolicitacaoFuncaoNFe").Value
    xLocal = 5
    CodigoEstabelecimento_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("CodigoEstabelecimento_MovSolicitacaoFuncaoNFe").Value
    
    xLocal = 6
    SerieNFe_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("SerieNFe_MovSolicitacaoFuncaoNFe").Value
    xLocal = 7
    NumeroNFe_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("NumeroNFe_MovSolicitacaoFuncaoNFe").Value
    xLocal = 8
    ChaveAcessoNFe_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("ChaveAcessoNFe_MovSolicitacaoFuncaoNFe").Value
    xLocal = 9
    IPComputadorAC_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("IPComputadorAC_MovSolicitacaoFuncaoNFe").Value
    xLocal = 10
    IPInternetAC_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("IPInternetAC_MovSolicitacaoFuncaoNFe").Value
    xLocal = 11
    
    SegurancaEstabelecimento_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("SegurancaEstabelecimento_MovSolicitacaoFuncaoNFe").Value
    xLocal = 12
    CodigoUsuario_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("CodigoUsuario_MovSolicitacaoFuncaoNFe").Value
    xLocal = 13
    VersaoAC_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("VersaoAC_MovSolicitacaoFuncaoNFe").Value
    xLocal = 14
    VersaoHost_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("VersaoHost_MovSolicitacaoFuncaoNFe").Value
    xLocal = 15
    Texto_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("Texto_MovSolicitacaoFuncaoNFe").Value
    xLocal = 16
    
    If IsNull(rsSolicitacaoFuncaoNFe("HoraAnalise_MovSolicitacaoFuncaoNFe").Value) Then
        HoraAnalise_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    Else
        HoraAnalise_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("HoraAnalise_MovSolicitacaoFuncaoNFe").Value
    End If
    xLocal = 17
    If IsNull(rsSolicitacaoFuncaoNFe("HoraAprovacao_MovSolicitacaoFuncaoNFe").Value) Then
        HoraAprovacao_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    Else
        HoraAprovacao_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("HoraAprovacao_MovSolicitacaoFuncaoNFe").Value
    End If
    xLocal = 18
    If IsNull(rsSolicitacaoFuncaoNFe("HoraCancelamentoHost_MovSolicitacaoFuncaoNFe").Value) Then
        HoraCancelamentoHost_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    Else
        HoraCancelamentoHost_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("HoraCancelamentoHost_MovSolicitacaoFuncaoNFe").Value
    End If
    xLocal = 19
    If IsNull(rsSolicitacaoFuncaoNFe("HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe").Value) Then
        HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    Else
        HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe").Value
    End If
    xLocal = 20
    If IsNull(rsSolicitacaoFuncaoNFe("HoraCancelamentoAC_MovSolicitacaoFuncaoNFe").Value) Then
        HoraCancelamentoAC_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    Else
        HoraCancelamentoAC_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("HoraCancelamentoAC_MovSolicitacaoFuncaoNFe").Value
    End If
    xLocal = 21
    Mensagem_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("Mensagem_MovSolicitacaoFuncaoNFe").Value
    xLocal = 22
    CodigoRetorno_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("CodigoRetorno_MovSolicitacaoFuncaoNFe").Value
    xLocal = 23
    NumeroLote_MovSolicitacaoFuncaoNFe = rsSolicitacaoFuncaoNFe("NumeroLote_MovSolicitacaoFuncaoNFe").Value
    xLocal = 24
    Exit Sub
    
trata_erro:
    MsgBox Err.Number & " - xLocal: " & xLocal & " - " & Err.Description
    
End Sub
    
    
