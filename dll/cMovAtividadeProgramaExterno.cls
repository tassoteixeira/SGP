VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovAtividadeProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    
    
    Private mDataHora_MovAtividadeProgramaExterno As Date
    Private mTipo_MovAtividadeProgramaExterno As String
    Private mIdEstabelecimento_MovAtividadeProgramaExterno As Integer
    Private mIpComputadorAC_MovAtividadeProgramaExterno As String
    Private mVersaoHost_MovAtividadeProgramaExterno As String
    Private mNomePrograma_MovAtividadeProgramaExterno As String
    Private mObservacao_MovAtividadeProgramaExterno As String

    Private rsMovAtividade As New adodb.Recordset

    Public Property Let DataHora_MovAtividadeProgramaExterno(ByVal Valor As Date)
        mDataHora_MovAtividadeProgramaExterno = Valor
    End Property
    Public Property Get DataHora_MovAtividadeProgramaExterno() As Date
        DataHora_MovAtividadeProgramaExterno = mDataHora_MovAtividadeProgramaExterno
    End Property
    
    Public Property Let Tipo_MovAtividadeProgramaExterno(ByVal Valor As String)
        mTipo_MovAtividadeProgramaExterno = Valor
    End Property
    Public Property Get Tipo_MovAtividadeProgramaExterno() As String
        Tipo_MovAtividadeProgramaExterno = mTipo_MovAtividadeProgramaExterno
    End Property
    
    Public Property Let IdEstabelecimento_MovAtividadeProgramaExterno(ByVal Valor As Integer)
        mIdEstabelecimento_MovAtividadeProgramaExterno = Valor
    End Property
    Public Property Get IdEstabelecimento_MovAtividadeProgramaExterno() As Integer
        IdEstabelecimento_MovAtividadeProgramaExterno = mIdEstabelecimento_MovAtividadeProgramaExterno
    End Property

    Public Property Let IpComputadorAC_MovAtividadeProgramaExterno(ByVal Valor As String)
        mIpComputadorAC_MovAtividadeProgramaExterno = Valor
    End Property
    Public Property Get IpComputadorAC_MovAtividadeProgramaExterno() As String
        IpComputadorAC_MovAtividadeProgramaExterno = mIpComputadorAC_MovAtividadeProgramaExterno
    End Property

    Public Property Let VersaoHost_MovAtividadeProgramaExterno(ByVal Valor As String)
        mVersaoHost_MovAtividadeProgramaExterno = Valor
    End Property
    Public Property Get VersaoHost_MovAtividadeProgramaExterno() As String
        VersaoHost_MovAtividadeProgramaExterno = mVersaoHost_MovAtividadeProgramaExterno
    End Property
    
    Public Property Let NomePrograma_MovAtividadeProgramaExterno(ByVal Valor As String)
        mNomePrograma_MovAtividadeProgramaExterno = Valor
    End Property
    Public Property Get NomePrograma_MovAtividadeProgramaExterno() As String
        NomePrograma_MovAtividadeProgramaExterno = mNomePrograma_MovAtividadeProgramaExterno
    End Property
    
    Public Property Let Observacao_MovAtividadeProgramaExterno(ByVal Valor As String)
        mObservacao_MovAtividadeProgramaExterno = Valor
    End Property
    Public Property Get Observacao_MovAtividadeProgramaExterno() As String
        Observacao_MovAtividadeProgramaExterno = mObservacao_MovAtividadeProgramaExterno
    End Property
    '
    
    
Public Function Incluir() As Boolean
    
Dim lRecordsAffected As Long
On Error GoTo Err_Incluir

        Incluir = False
        gSQL = "INSERT INTO MovAtividadeProgramaExterno(DataHora_MovAtividadeProgramaExterno,Tipo_MovAtividadeProgramaExterno,IdEstabelecimento_MovAtividadeProgramaExterno,"
        gSQL = gSQL & "IpComputadorAC_MovAtividadeProgramaExterno,VersaoHost_MovAtividadeProgramaExterno,NomePrograma_MovAtividadeProgramaExterno,"
        gSQL = gSQL & "Observacao_MovAtividadeProgramaExterno)VALUES ( "
        Call sqlDataHora(1, mDataHora_MovAtividadeProgramaExterno, ", ")
        Call sqlTexto(1, mTipo_MovAtividadeProgramaExterno, ", ")
        Call sqlNumero(1, mIdEstabelecimento_MovAtividadeProgramaExterno, ", ")
        Call sqlTexto(1, mIpComputadorAC_MovAtividadeProgramaExterno, ", ")
        Call sqlTexto(1, mVersaoHost_MovAtividadeProgramaExterno, ", ")
        Call sqlTexto(1, mNomePrograma_MovAtividadeProgramaExterno, ", ")
        Call sqlTexto(1, mObservacao_MovAtividadeProgramaExterno, " )")

        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords

        
        If lRecordsAffected > 0 Then
            Incluir = True
        End If
        Exit Function
        
Err_Incluir:
        

End Function



Public Function LocalizarUltimoTipoPrograma(ByVal pTipo As String, ByVal pIpComputador As String, ByVal pNomePrograma As String) As Boolean
        Dim xCondicao As String
        xCondicao = ""
        
On Error GoTo trata_erro

        
        LocalizarUltimoTipoPrograma = False
        xCondicao = xCondicao & " WHERE NomePrograma_MovAtividadeProgramaExterno = " & preparaTexto(pNomePrograma)
        xCondicao = xCondicao & " AND IpComputadorAC_MovAtividadeProgramaExterno = " & preparaTexto(pIpComputador)
        
        If Trim(pTipo) <> "" Then
            xCondicao = xCondicao & " AND Tipo_MovAtividadeProgramaExterno = " & preparaTexto(pTipo)
        End If
        
        
        gSQL = PreparaSQL(xCondicao, "ORDER BY DataHora_MovAtividadeProgramaExterno DESC")
        
        LocalizarUltimoTipoPrograma = Localizar(1)
        If LocalizarUltimoTipoPrograma Then
            AtribuiValor
        End If
        
        Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description

End Function
    Private Sub AtribuiValor()
        Dim xLocal As Integer

        xLocal = 1
        mDataHora_MovAtividadeProgramaExterno = rsMovAtividade("DataHora_MovAtividadeProgramaExterno").Value
        xLocal = 2
        mTipo_MovAtividadeProgramaExterno = rsMovAtividade("Tipo_MovAtividadeProgramaExterno").Value
        xLocal = 3
        mIdEstabelecimento_MovAtividadeProgramaExterno = rsMovAtividade("IdEstabelecimento_MovAtividadeProgramaExterno").Value
        xLocal = 4
        mIpComputadorAC_MovAtividadeProgramaExterno = rsMovAtividade("IpComputadorAC_MovAtividadeProgramaExterno").Value
        xLocal = 5
        mVersaoHost_MovAtividadeProgramaExterno = rsMovAtividade("VersaoHost_MovAtividadeProgramaExterno").Value
        xLocal = 6
        mNomePrograma_MovAtividadeProgramaExterno = rsMovAtividade("NomePrograma_MovAtividadeProgramaExterno").Value
        xLocal = 7
        mObservacao_MovAtividadeProgramaExterno = rsMovAtividade("Observacao_MovAtividadeProgramaExterno").Value


    End Sub


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovAtividade = New adodb.Recordset
    rsMovAtividade.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovAtividade.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovAtividade.EOF Then
        AtualizaRecordset = True
    End If
End Function
    
Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
        Localizar = False
        If AtualizaRecordset(pQtdRegistro) Then
            Localizar = True
        End If
End Function

Private Function PreparaSQL(ByVal pCondicao As String, ByVal pTipoOrdem As String) As String

        Dim xSQL As String
        xSQL = " SELECT DataHora_MovAtividadeProgramaExterno,Tipo_MovAtividadeProgramaExterno,IdEstabelecimento_MovAtividadeProgramaExterno,"
        xSQL = xSQL & "IpComputadorAC_MovAtividadeProgramaExterno,VersaoHost_MovAtividadeProgramaExterno,NomePrograma_MovAtividadeProgramaExterno,"
        xSQL = xSQL & "Observacao_MovAtividadeProgramaExterno FROM MovAtividadeProgramaExterno"


        PreparaSQL = Trim(xSQL & " " & pCondicao & " " & pTipoOrdem)
End Function

