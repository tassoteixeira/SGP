VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSolicitacaoFuncaoNFe"
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
    Private mCodigoEstabelecimento As Integer
    Private mSerieNFe As String
    Private mNumeroNFe As Long
    Private mChaveAcessoNFe As String
    Private mIPComputadorAC As String
    Private mIPInternetAC As String
    Private mSegurancaEstabelecimento As String
    Private mCodigoUsuario As Integer
    Private mVersaoAC As String
    Private mVersaoHost As String
    Private mTexto As String
    Private mHoraAnalise As Date
    Private mHoraAprovacao As Date
    Private mHoraCancelamentoHost As Date
    Private mHoraConfirmacaoAC As Date
    Private mHoraCancelamentoAC As Date
    Private mMensagem As String

    Private rsSolicitacaoFuncaoNFe As New adodb.Recordset

    Public Property Let NSU(ByVal Valor As Long)
        mNSU = Valor
    End Property
    Public Property Get NSU() As Long
        NSU = mNSU
    End Property
    
    Public Property Let NumeroControleSolicitacao(ByVal Valor As Long)
        mNumeroControleSolicitacao = Valor
    End Property
    Public Property Get NumeroControleSolicitacao() As Long
        NumeroControleSolicitacao = mNumeroControleSolicitacao
    End Property
    
    Public Property Let DataSolicitacao(ByVal Valor As Date)
        mDataSolicitacao = Valor
    End Property
    Public Property Get DataSolicitacao() As Date
        DataSolicitacao = mDataSolicitacao
    End Property
    
    Public Property Let HoraSolicitacao(ByVal Valor As Date)
        mHoraSolicitacao = Valor
    End Property
    Public Property Get HoraSolicitacao() As Date
        HoraSolicitacao = mHoraSolicitacao
    End Property
    
    Public Property Let TipoOperacao(ByVal Valor As String)
        mTipoOperacao = Valor
    End Property
    Public Property Get TipoOperacao() As String
        TipoOperacao = mTipoOperacao
    End Property
    
    Public Property Let CodigoEstabelecimento(ByVal Valor As Integer)
        mCodigoEstabelecimento = Valor
    End Property
    Public Property Get CodigoEstabelecimento() As Integer
        CodigoEstabelecimento = mCodigoEstabelecimento
    End Property
    
    Public Property Let SerieNFe(ByVal Valor As String)
        mSerieNFe = Valor
    End Property
    Public Property Get SerieNFe() As String
        SerieNFe = mSerieNFe
    End Property
    
    Public Property Let NumeroNFe(ByVal Valor As Long)
        mNumeroNFe = Valor
    End Property
    Public Property Get NumeroNFe() As Long
        NumeroNFe = mNumeroNFe
    End Property
    
    Public Property Let ChaveAcessoNFe(ByVal Valor As String)
        mChaveAcessoNFe = Valor
    End Property
    Public Property Get ChaveAcessoNFe() As String
        ChaveAcessoNFe = mChaveAcessoNFe
    End Property
    
    Public Property Let IPComputadorAC(ByVal Valor As String)
        mIPComputadorAC = Valor
    End Property
    Public Property Get IPComputadorAC() As String
        IPComputadorAC = mIPComputadorAC
    End Property
    
    Public Property Let IPInternetAC(ByVal Valor As String)
        mIPInternetAC = Valor
    End Property
    Public Property Get IPInternetAC() As String
        IPInternetAC = mIPInternetAC
    End Property
    
    Public Property Let SegurancaEstabelecimento(ByVal Valor As String)
        mSegurancaEstabelecimento = Valor
    End Property
    Public Property Get SegurancaEstabelecimento() As String
        SegurancaEstabelecimento = mSegurancaEstabelecimento
    End Property
    
    Public Property Let CodigoUsuario(ByVal Valor As Integer)
        mCodigoUsuario = Valor
    End Property
    Public Property Get CodigoUsuario() As Integer
        CodigoUsuario = mCodigoUsuario
    End Property
    
    Public Property Let VersaoAC(ByVal Valor As String)
        mVersaoAC = Valor
    End Property
    Public Property Get VersaoAC() As String
        VersaoAC = mVersaoAC
    End Property
    
    Public Property Let VersaoHost(ByVal Valor As String)
        mVersaoHost = Valor
    End Property
    Public Property Get VersaoHost() As String
        VersaoHost = mVersaoHost
    End Property
    
    Public Property Let Texto(ByVal Valor As String)
        mTexto = Valor
    End Property
    Public Property Get Texto() As String
        Texto = mTexto
    End Property
    
    Public Property Let HoraAnalise(ByVal Valor As Date)
        mHoraAnalise = Valor
    End Property
    Public Property Get HoraAnalise() As Date
        HoraAnalise = mHoraAnalise
    End Property
    
    Public Property Let HoraAprovacao(ByVal Valor As Date)
        mHoraAprovacao = Valor
    End Property
    Public Property Get HoraAprovacao() As Date
        HoraAprovacao = mHoraAprovacao
    End Property
    
    Public Property Let HoraCancelamentoHost(ByVal Valor As Date)
        mHoraCancelamentoHost = Valor
    End Property
    Public Property Get HoraCancelamentoHost() As Date
        HoraCancelamentoHost = mHoraCancelamentoHost
    End Property
    
    Public Property Let HoraConfirmacaoAC(ByVal Valor As Date)
        mHoraConfirmacaoAC = Valor
    End Property
    Public Property Get HoraConfirmacaoAC() As Date
        HoraConfirmacaoAC = mHoraConfirmacaoAC
    End Property
    
    Public Property Let HoraCancelamentoAC(ByVal Valor As Date)
        mHoraCancelamentoAC = Valor
    End Property
    Public Property Get HoraCancelamentoAC() As Date
        HoraCancelamentoAC = mHoraCancelamentoAC
    End Property
    
    Public Property Let Mensagem(ByVal Valor As String)
        mMensagem = Valor
    End Property
    Public Property Get Mensagem() As String
        Mensagem = mMensagem
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
        PreparaSQL = Trim("SELECT NSU,[Numero de Controle de Solicitacao],[Data da Solicitacao],[Hora da Solicitacao],[Tipo de Operacao],[Codigo do Estabelecimento],[Serie da NFe],[Numero da NFe],[Chave de Acesso da NFe],[IP Computador AC],[IP Internet AC],[Seguranca do Estabelecimento],[Codigo do Usuario],[Versao AC],[Versao Host], Texto, [Hora da Analise],[Hora da Aprovacao],[Hora Cancelamento Host],[Hora da Confirmacao AC],[Hora do Cancelamento AC], Mensagem FROM SolicitacaoFuncaoNFe" & " " & xCondicao & " " & xTipoOrdem)
    End Function

    
Public Function BuscaMensagem(ByVal pNSU As Integer) As String
      
      On Error GoTo trata_erro
            
            BuscaMensagem = ""
            gSQL = ""
            
            gSQL = "SELECT Mensagem"
            gSQL = gSQL & "  FROM SolicitacaoFuncaoNFe"
            gSQL = gSQL & " WHERE NSU = " & pNSU
            
            If Localizar(1) Then
                BuscaMensagem = rsSolicitacaoFuncaoNFe("Mensagem")
            End If
            rsSolicitacaoFuncaoNFe.Close
            Set rsSolicitacaoFuncaoNFe = Nothing
            Exit Function
        
trata_erro:
            MsgBox Err.Number & " - " & Err.Description
    
End Function
Public Function DefineHoraAnalise(ByVal pNSU As Integer, ByVal pHoraAnalise As Date, ByVal pVersaoHost As String) As Boolean
        
        
Dim lRecordsAffected As Long
    
On Error GoTo Err_DefineHoraAnalise
    
        DefineHoraAnalise = False
        gSQL = "UPDATE SolicitacaoFuncaoNFe SET "
        Call sqlHora(2, "[Hora da Analise] = ", pHoraAnalise)
        Call sqlTexto(2, ", [Versao Host] = ", pVersaoHost)
        gSQL = gSQL & " WHERE NSU = " & pNSU
        gSQL = gSQL & " AND [Hora da Analise] IS NULL"
        
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        
        If lRecordsAffected > 0 Then
            DefineHoraAnalise = True
        End If
        Exit Function
    
Err_DefineHoraAnalise:

End Function
Public Function DefineHoraAprovacao(ByVal pNSU As Integer, ByVal pHoraAprovacao As Date, ByVal pVersaoHost As String) As Boolean
        
Dim lRecordsAffected As Long
    
On Error GoTo Err_DefineHoraAprovacao
    
        DefineHoraAprovacao = False
        gSQL = "UPDATE SolicitacaoFuncaoNFe SET "
        Call sqlHora(2, "[Hora da Aprovacao] = ", pHoraAprovacao)
        Call sqlTexto(2, ", [Versao Host] = ", pVersaoHost)
        gSQL = gSQL & " WHERE NSU = " & pNSU
        gSQL = gSQL & " AND [Hora da Aprovacao] IS NULL"
        
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        
        If lRecordsAffected > 0 Then
            DefineHoraAprovacao = True
        End If
        Exit Function
    
Err_DefineHoraAprovacao:

End Function
    
Public Function DefineHoraCancelamentoAC(ByVal pNSU As Integer, ByVal pHoraCancelamentoAC As Date) As Boolean
        
        
Dim lRecordsAffected As Long
    
On Error GoTo Err_DefineHoraCancelamentoAC
    
        DefineHoraCancelamentoAC = False
        gSQL = "UPDATE SolicitacaoFuncaoNFe SET "
        Call sqlHora(2, "[Hora do Cancelamento AC] = ", pHoraCancelamentoAC)
        gSQL = gSQL & " WHERE NSU = " & pNSU
        
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        
        If lRecordsAffected > 0 Then
            DefineHoraCancelamentoAC = True
        End If
        Exit Function
    
Err_DefineHoraCancelamentoAC:
        
End Function
Public Function DefineHoraCancelamentoHost(ByVal pNSU As Integer, ByVal pHoraCancelamentoHost As Date, ByVal pVersaoHost As String) As Boolean
        
        Dim lRecordsAffected As Long
    
On Error GoTo Err_DefineHoraCancelamentoHost
    
        DefineHoraCancelamentoHost = False
        gSQL = "UPDATE SolicitacaoFuncaoNFe SET "
        Call sqlHora(2, "[Hora Cancelamento Host] = ", pHoraCancelamentoHost)
        Call sqlTexto(2, ", [Versao Host] = ", pVersaoHost)
        gSQL = gSQL & " WHERE NSU = " & pNSU
        gSQL = gSQL & " AND [Hora Cancelamento Host] IS NULL"
        
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        
        If lRecordsAffected > 0 Then
            DefineHoraCancelamentoHost = True
        End If
        Exit Function
    
Err_DefineHoraCancelamentoHost:

End Function
Public Function DefineHoraConfirmacaoAC(ByVal pNSU As Integer, ByVal pHoraConfirmacaoAC As Date) As Boolean
        
        
Dim lRecordsAffected As Long
    
On Error GoTo Err_DefineHoraConfirmacaoAC
    
        DefineHoraConfirmacaoAC = False
        gSQL = "UPDATE SolicitacaoFuncaoNFe SET "
        Call sqlHora(2, "[Hora da Confirmacao AC] = ", pHoraConfirmacaoAC)
        gSQL = gSQL & " WHERE NSU = " & pNSU
        gSQL = gSQL & " AND [Hora da Confirmacao AC] IS NULL"

        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        
        If lRecordsAffected > 0 Then
            DefineHoraConfirmacaoAC = True
        End If
        Exit Function
    
Err_DefineHoraConfirmacaoAC:

End Function
Public Function DefineMensagem(ByVal pNSU As Integer, ByVal pMensagem As String) As Boolean
        
Dim lRecordsAffected As Long
    
On Error GoTo Err_DefineMensagem
    
        DefineMensagem = False
        gSQL = "UPDATE SolicitacaoFuncaoNFe SET "
        Call sqlHora(2, "Mensagem = ", pMensagem)
        gSQL = gSQL & " WHERE NSU = " & pNSU

        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        
        If lRecordsAffected > 0 Then
            DefineMensagem = True
        End If
        Exit Function
    
Err_DefineMensagem:

        
End Function
Public Function Alterar(ByVal pNSU As Integer) As Boolean

        Dim lRecordsAffected As Long
        
        On Error GoTo Err_Alterar
            Alterar = False
            gSQL = "UPDATE SolicitacaoFuncaoNFe SET "
            Call sqlNumero(2, "NSU = ", CStr(mNSU))
            Call sqlNumero(2, ", [Numero de Controle de Solicitacao] = ", CStr(mNumeroControleSolicitacao))
            Call sqlData(2, ", [Data da Solicitacao] = ", CStr(mDataSolicitacao))
            Call sqlHora(2, ", [Hora da Solicitacao] = ", CStr(mHoraSolicitacao))
            Call sqlTexto(2, ", [Tipo de Operacao] = ", mTipoOperacao)
            Call sqlNumero(2, ", [Codigo do Estabelecimento] = ", CStr(mCodigoEstabelecimento))
            Call sqlTexto(2, ", [Serie da NFe] = ", mSerieNFe)
            Call sqlNumero(2, ", [Numero da NFe] = ", CStr(mNumeroNFe))
            Call sqlTexto(2, ", [Chave de Acesso da NFe] = ", mChaveAcessoNFe)
            Call sqlTexto(2, ", [IP Computador AC] = ", mIPComputadorAC)
            Call sqlTexto(2, ", [IP Internet AC] = ", mIPInternetAC)
            Call sqlTexto(2, ", [Seguranca do Estabelecimento] = ", mSegurancaEstabelecimento)
            Call sqlNumero(2, ", [Codigo do Usuario] = ", CStr(mCodigoUsuario))
            Call sqlTexto(2, ", [Versao AC] = ", mVersaoAC)
            Call sqlTexto(2, ", [Versao Host] = ", mVersaoHost)
            Call sqlTexto(2, ", Texto = ", mTexto)
            Call sqlHora(2, ", [Hora da Analise] = ", CStr(mHoraAnalise))
            Call sqlHora(2, ", [Hora da Aprovacao] = ", CStr(mHoraAprovacao))
            Call sqlHora(2, ", [Hora Cancelamento Host] = ", CStr(mHoraCancelamentoHost))
            Call sqlHora(2, ", [Hora da Confirmacao AC] = ", CStr(mHoraConfirmacaoAC))
            Call sqlHora(2, ", [Hora do Cancelamento AC] = ", CStr(mHoraCancelamentoAC))
            Call sqlTexto(2, ", Mensagem = ", mMensagem)
            gSQL = gSQL & " WHERE NSU = " & pNSU
            
            gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
            
            If lRecordsAffected > 0 Then
                Alterar = True
            End If
            
Err_Alterar:
            
End Function
Public Function Excluir(ByVal pNSU As Integer) As Boolean
    
    
    Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM SolicitacaoFuncaoNFe"
    gSQL = gSQL & " WHERE NSU = " & pNSU
    
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
        mNSU = ProximoRegistroNFCE()
        
        gSQL = "INSERT INTO SolicitacaoFuncaoNFe ( NSU, [Numero de Controle de Solicitacao], [Data da Solicitacao], [Hora da Solicitacao], [Tipo de Operacao], "
        gSQL = gSQL & "[Codigo do Estabelecimento], [Serie da NFe], [Numero da NFe], [Chave de Acesso da NFe], [IP Computador AC], "
        gSQL = gSQL & "[IP Internet AC], [Seguranca do Estabelecimento], [Codigo do Usuario], [Versao AC], [Versao Host], Texto, "
        gSQL = gSQL & "[Hora da Analise], [Hora da Aprovacao], [Hora Cancelamento Host], [Hora da Confirmacao AC], [Hora do Cancelamento AC], "
        gSQL = gSQL & "Mensagem ) VALUES ( "
        Call sqlNumero(1, CStr(mNSU), ", ")
        Call sqlNumero(1, CStr(mNumeroControleSolicitacao), ", ")
        Call sqlData(1, CStr(mDataSolicitacao), ", ")
        Call sqlHora(1, CStr(mHoraSolicitacao), ", ")
        Call sqlTexto(1, mTipoOperacao, ", ")
        Call sqlNumero(1, CStr(mCodigoEstabelecimento), ", ")
        Call sqlTexto(1, mSerieNFe, ", ")
        Call sqlTexto(1, mNumeroNFe, ", ")
        Call sqlTexto(1, mChaveAcessoNFe, ", ")
        Call sqlTexto(1, mIPComputadorAC, ", ")
        Call sqlTexto(1, mIPInternetAC, ", ")
        Call sqlTexto(1, mSegurancaEstabelecimento, ", ")
        Call sqlNumero(1, CStr(mCodigoUsuario), ", ")
        Call sqlTexto(1, mVersaoAC, ", ")
        Call sqlTexto(1, mVersaoHost, ", ")
        Call sqlTexto(1, mTexto, ", ")
        Call sqlHora(1, CStr(mHoraAnalise), ", ")
        Call sqlHora(1, CStr(mHoraAprovacao), ", ")
        Call sqlHora(1, CStr(mHoraCancelamentoHost), ", ")
        Call sqlHora(1, CStr(mHoraConfirmacaoAC), ", ")
        Call sqlHora(1, CStr(mHoraCancelamentoAC), ", ")
        Call sqlTexto(1, mMensagem, " )")
        
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
            
        If lRecordsAffected > 0 Then
           Incluir = True
        End If
        
Err_Incluir:

End Function
Public Function ProximoRegistroNFCE() As Integer
    
    
     On Error GoTo trata_erro
            
            ProximoRegistroNFCE = 1
            gSQL = ""
            
            gSQL = "SELECT NSU"
            gSQL = gSQL & "  FROM SolicitacaoFuncaoNFe"
            gSQL = gSQL & " ORDER BY NSU DESC"
            If Localizar(1) Then
                ProximoRegistroNFCE = rsSolicitacaoFuncaoNFe("NSU") + 1
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
            
            gSQL = "SELECT [Numero da NFe]"
            gSQL = gSQL & "  FROM SolicitacaoFuncaoNFe"
            gSQL = gSQL & "  WHERE [Tipo de Operacao] = " & preparaTexto("NFCe 3.10")
            gSQL = gSQL & " ORDER BY NSU DESC"
            If Localizar(1) Then
                ProximoNumeroNFCE = rsSolicitacaoFuncaoNFe("Numero da NFe") + 1
            End If
            rsSolicitacaoFuncaoNFe.Close
            Set rsSolicitacaoFuncaoNFe = Nothing
            Exit Function
        
trata_erro:
            MsgBox Err.Number & " - " & Err.Description
    
End Function


Private Sub AtribuiValor()
        Dim xLocal As Integer
            
            xLocal = 1
            NSU = rsSolicitacaoFuncaoNFe("NSU").Value
            xLocal = 2
            NumeroControleSolicitacao = rsSolicitacaoFuncaoNFe("Numero de Controle de Solicitacao").Value
            xLocal = 3
            DataSolicitacao = rsSolicitacaoFuncaoNFe("Data da Solicitacao").Value
            xLocal = 4
            HoraSolicitacao = rsSolicitacaoFuncaoNFe("Hora da Solicitacao").Value
            xLocal = 5
            TipoOperacao = rsSolicitacaoFuncaoNFe("Tipo de Operacao").Value
            xLocal = 6
            CodigoEstabelecimento = rsSolicitacaoFuncaoNFe("Codigo do Estabelecimento").Value
            xLocal = 7
            SerieNFe = rsSolicitacaoFuncaoNFe("Serie da NFe").Value
            xLocal = 8
            NumeroNFe = rsSolicitacaoFuncaoNFe("Numero da NFe").Value
            xLocal = 9
            ChaveAcessoNFe = rsSolicitacaoFuncaoNFe("Chave de Acesso da NFe").Value
            xLocal = 10
            IPComputadorAC = rsSolicitacaoFuncaoNFe("IP Computador AC").Value
            xLocal = 11
            IPInternetAC = rsSolicitacaoFuncaoNFe("IP Internet AC").Value
            xLocal = 12
            SegurancaEstabelecimento = rsSolicitacaoFuncaoNFe("Seguranca do Estabelecimento").Value
            xLocal = 13
            CodigoUsuario = rsSolicitacaoFuncaoNFe("Codigo do Usuario").Value
            xLocal = 14
            VersaoAC = rsSolicitacaoFuncaoNFe("Versao AC").Value
            xLocal = 15
            VersaoHost = rsSolicitacaoFuncaoNFe("Versao Host").Value
            xLocal = 16
            Texto = rsSolicitacaoFuncaoNFe("Texto").Value
            xLocal = 17
            If IsNull(rsSolicitacaoFuncaoNFe("Hora da Analise").Value) Then
                HoraAnalise = CDate("00:00:00")
            Else
                HoraAnalise = rsSolicitacaoFuncaoNFe("Hora da Analise").Value
            End If
            xLocal = 18
            If IsNull(rsSolicitacaoFuncaoNFe("Hora da Aprovacao").Value) Then
                HoraAprovacao = CDate("00:00:00")
            Else
                HoraAprovacao = rsSolicitacaoFuncaoNFe("Hora da Aprovacao").Value
            End If
            xLocal = 19
            If IsNull(rsSolicitacaoFuncaoNFe("Hora Cancelamento Host").Value) Then
                HoraCancelamentoHost = CDate("00:00:00")
            Else
                HoraCancelamentoHost = rsSolicitacaoFuncaoNFe("Hora Cancelamento Host").Value
            End If
            xLocal = 20
            If IsNull(rsSolicitacaoFuncaoNFe("Hora da Confirmacao AC").Value) Then
                HoraConfirmacaoAC = CDate("00:00:00")
            Else
                HoraConfirmacaoAC = rsSolicitacaoFuncaoNFe("Hora da Confirmacao AC").Value
            End If
            xLocal = 21
            If IsNull(rsSolicitacaoFuncaoNFe("Hora do Cancelamento AC").Value) Then
                HoraCancelamentoAC = CDate("00:00:00")
            Else
                HoraCancelamentoAC = rsSolicitacaoFuncaoNFe("Hora do Cancelamento AC").Value
            End If
            xLocal = 22
            Mensagem = rsSolicitacaoFuncaoNFe("Mensagem").Value
    
End Sub
    
    
    
    
    

'
'
'    Public Function LocalizarAnaliseEOuAprovacao() As Boolean
'        LocalizarAnaliseEOuAprovacao = False
'        gSQL = "SELECT NSU"
'        gSQL += "  FROM SolicitacaoFuncaoNFe"
'        gSQL += " WHERE [Hora do Cancelamento AC] IS NULL"
'        gSQL += "   AND [Hora da Confirmacao AC] IS NULL"
'        gSQL += "   AND [Hora Cancelamento Host] IS NULL"
'        gSQL += "   AND ([Hora da Analise] IS NULL"
'        gSQL += "    OR [Hora da Aprovacao] IS NULL)"
'        gSQL += " ORDER BY [Data da Solicitacao] ASC, [Hora da Solicitacao] ASC"
'        LocalizarAnaliseEOuAprovacao = Localizar(1, True, False, True)
'    End Function
'    Public Function LocalizarCodigo(ByVal pNSU As Integer) As Boolean
'        Dim xCondicao As String
'
'        LocalizarCodigo = False
'        xCondicao = " WHERE NSU = " & pNSU
'        gSQL = PreparaSQL(xCondicao, "").Value
'        LocalizarCodigo = Localizar(1, True, True, True)
'    End Function
'    'Public Function MensagemAtual(ByVal pNSU As Integer) As String
'    '    MensagemAtual = ""
'    '    Try
'    '        gSQL = "SELECT Mensagem"
'    '        gSQL += "  FROM SolicitacaoFuncaoNFe"
'    '        gSQL += " WHERE NSU = " & pNSU
'    '        gSQL += " ORDER BY NSU DESC"
'    '        If Localizar(1, True, False, False) Then
'    '            MensagemAtual = rsSolicitacaoFuncaoNFe("Mensagem").Value
'    '        End If
'    '        drTabela.Close()
'    '        cmd.Dispose()
'    '    Catch ex As Exception
'    '        CriaLogRN(Me.GetType.Name & ":MensagemAtual - Erro inesperado", Err.Description, gSQL)
'    '    End Try
'    'End Function
'    Public Function MontaDS(ByVal pSQL As String) As DataSet
'        Dim dsTabela As New DataSet
'
'        Try
'            daTabela = New OleDbDataAdapter(pSQL, gConnNFe)
'            daTabela.Fill(dsTabela, "SolicitacaoFuncaoNFe").Value
'            Return dsTabela
'        Catch
'            CriaLogRN(Me.GetType.Name & ":MontaDS - Erro não identificado.", Err.Description, pSQL)
'            dsTabela = New DataSet
'            Return dsTabela
'        Finally
'        End Try
'    End Function
'    Public Function VerificaSeEstaAprovado(ByVal pNSU As Integer) As Boolean
'        VerificaSeEstaAprovado = False
'        gSQL = "SELECT NSU"
'        gSQL += "  FROM SolicitacaoFuncaoNFe"
'        gSQL += " WHERE NSU = " & pNSU
'        gSQL += "   AND NOT [Hora da Aprovacao] IS NULL"
'        VerificaSeEstaAprovado = Localizar(1, True, False, True)
'    End Function
'    Public Function VerificaSeEstaEmAnalise(ByVal pNSU As Integer) As Boolean
'        VerificaSeEstaEmAnalise = False
'        gSQL = "SELECT NSU"
'        gSQL += "  FROM SolicitacaoFuncaoNFe"
'        gSQL += " WHERE NSU = " & pNSU
'        gSQL += "   AND NOT [Hora da Analise] IS NULL"
'        VerificaSeEstaEmAnalise = Localizar(1, True, False, True)
'    End Function
'    Public Function VerificaSeEstaCanceladoHost(ByVal pNSU As Integer) As Boolean
'        VerificaSeEstaCanceladoHost = False
'        gSQL = "SELECT NSU"
'        gSQL += "  FROM SolicitacaoFuncaoNFe"
'        gSQL += " WHERE NSU = " & pNSU
'        gSQL += "   AND NOT [Hora Cancelamento Host] IS NULL"
'        VerificaSeEstaCanceladoHost = Localizar(1, True, False, True)
'    End Function
'
'#End Region
'
'#Region " Funções/Procedures Internas da Classe "
'
'
'    Private Function Localizar_OLD(ByVal pQtdRegistro As Short, ByVal pLeRegistro As Boolean, ByVal pAtribuiValor As Boolean, ByVal pFecha As Boolean) As Boolean
'        Dim i As Short
'
'        Localizar = False
'        i = Len(gSQL)
'        If pQtdRegistro > 0 Then
'            gSQL = Mid(gSQL, 1, 6) & " Top " & pQtdRegistro & Mid(gSQL, 7, i - 6)
'        End If
'        Try
'            cmd = New OleDbCommand(gSQL, gConnNFe)
'            drTabela = cmd.ExecuteReader
'            If pLeRegistro = True Then
'                If drTabela.Read() Then
'                    Localizar = True
'                    If pAtribuiValor Then
'                        AtribuiValor()
'                    End If
'                End If
'            Else
'                Localizar = True
'            End If
'        Catch
'            CriaLogRN(Me.GetType.Name & ":Localizar - Erro não identificado.", Err.Description, gSQL)
'        Finally
'            If pFecha Then
'                drTabela.Close()
'                cmd.Dispose()
'            End If
'        End Try
'    End Function
'    Private Function PreparaSQL(ByVal pCondicao As String, ByVal pTipoOrdem As String) As String
'        PreparaSQL = Trim("SELECT NSU, [Numero de Controle de Solicitacao], [Data da Solicitacao], [Hora da Solicitacao], [Tipo de Operacao], [Codigo do Estabelecimento], [Serie da NFe], [Numero da NFe], [Chave de Acesso da NFe], [IP Computador AC], [IP Internet AC], [Seguranca do Estabelecimento], [Codigo do Usuario], [Versao AC], [Versao Host], Texto, [Hora da Analise], [Hora da Aprovacao], [Hora Cancelamento Host], [Hora da Confirmacao AC], [Hora do Cancelamento AC], Mensagem FROM SolicitacaoFuncaoNFe" & " " & pCondicao & " " & pTipoOrdem)
'    End Function

'#End Region
