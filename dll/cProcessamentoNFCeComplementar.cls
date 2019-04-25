VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProcessaNFCeComplementar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Dim mIdEstabelecimento As Integer
Dim mDataHoraInicio As Date
Dim mPDV As Integer
Dim mDataHoraEmissao As Date
Dim mDataHoraTermino As Date
Dim mObservacao As String

Private rsProcessa As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get IdEstabelecimento() As Integer
    IdEstabelecimento = mIdEstabelecimento
End Property
Public Property Let IdEstabelecimento(ByVal Valor As Integer)
    mIdEstabelecimento = Valor
End Property
Public Property Get DataHoraInicio() As Date
    DataHoraInicio = mDataHoraInicio
End Property
Public Property Let DataHoraInicio(ByVal Valor As Date)
    mDataHoraInicio = Valor
End Property
Public Property Get PDV() As Integer
    PDV = mPDV
End Property
Public Property Let PDV(ByVal Valor As Integer)
    mPDV = Valor
End Property
Public Property Get DataHoraEmissao() As Date
    DataHoraEmissao = mDataHoraEmissao
End Property
Public Property Let DataHoraEmissao(ByVal Valor As Date)
    mDataHoraEmissao = Valor
End Property
Public Property Get DataHoraTermino() As Date
    DataHoraTermino = mDataHoraTermino
End Property
Public Property Let DataHoraTermino(ByVal Valor As Date)
    mDataHoraTermino = Valor
End Property
Public Property Get Observacao() As String
    Observacao = mObservacao
End Property
Public Property Let Observacao(ByVal Valor As String)
    mObservacao = Valor
End Property



Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    

    gSQL = "INSERT INTO ProcessamentoNFCeComplementar (IdEstabelecimento_ProcessamentoNFCeComplementar, DataHoraInicio_ProcessamentoNFCeComplementar, PDV_ProcessamentoNFCeComplementar, "
    gSQL = gSQL & " DataHoraEmissao_ProcessamentoNFCeComplementar, DataHoraTermino_ProcessamentoNFCeComplementar, Observacao_ProcessamentoNFCeComplementar "
    gSQL = gSQL & ") VALUES ("
    Call sqlNumero(1, mIdEstabelecimento, ", ")
    Call sqlDataHora(1, mDataHoraInicio, ", ")
    Call sqlNumero(1, mPDV, ", ")
    Call sqlDataHora(1, mDataHoraEmissao, ", ")
    Call sqlDataHora(1, mDataHoraTermino, ", ")
    Call sqlTexto(1, mObservacao, " )")
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    
    If lRecordsAffected > 0 Then
       Incluir = True
    End If
    Exit Function
    
Err_Incluir:
End Function

Public Function Alterar(ByVal pIdEstabelecimento As Integer, ByVal pDataHoraInicio As Date, ByVal pPDV As Integer) As Boolean
    Dim lRecordsAffected As Long
    
    On Error GoTo Err_Alterar
    
    Alterar = False
    
    gSQL = "UPDATE ProcessamentoNFCeComplementar SET "
    Call sqlNumero(2, "IdEstabelecimento_ProcessamentoNFCeComplementar = ", mIdEstabelecimento)
    Call sqlDataHora(2, ", DataHoraInicio_ProcessamentoNFCeComplementar = ", mDataHoraInicio)
    Call sqlNumero(2, ", PDV_ProcessamentoNFCeComplementar = ", mPDV)
    Call sqlDataHora(2, ", DataHoraEmissao_ProcessamentoNFCeComplementar = ", mDataHoraEmissao)
    Call sqlDataHora(2, ", DataHoraTermino_ProcessamentoNFCeComplementar = ", mDataHoraTermino)
    Call sqlTexto(2, ", Observacao_ProcessamentoNFCeComplementar = ", mObservacao)
    
    gSQL = gSQL & " WHERE IdEstabelecimento_ProcessamentoNFCeComplementar = " & pIdEstabelecimento
    gSQL = gSQL & " AND DataHoraInicio_ProcessamentoNFCeComplementar = " & preparaDataHora(pDataHoraInicio)
    gSQL = gSQL & " AND PDV_ProcessamentoNFCeComplementar = " & pPDV
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function


Err_Alterar:

End Function

Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsProcessa = New adodb.Recordset
    rsProcessa.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsProcessa.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsProcessa.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()

    IdEstabelecimento = rsProcessa("IdEstabelecimento_ProcessamentoNFCeComplementar").Value
    DataHoraInicio = rsProcessa("DataHoraInicio_ProcessamentoNFCeComplementar").Value
    PDV = rsProcessa("PDV_ProcessamentoNFCeComplementar").Value
    DataHoraEmissao = rsProcessa("DataHoraEmissao_ProcessamentoNFCeComplementar").Value
    DataHoraTermino = rsProcessa("DataHoraTermino_ProcessamentoNFCeComplementar").Value
    Observacao = rsProcessa("Observacao_ProcessamentoNFCeComplementar").Value
End Sub

Public Function DefinirDataHoraTermino(ByVal pIdEstabelecimento As Integer, ByVal pDataHoraInicio As Date, ByVal pPDV As Integer, ByVal pObservacao As String) As Boolean

Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    DefinirDataHoraTermino = False
    gSQL = "UPDATE ProcessamentoNFCeComplementar SET "
    Call sqlDataHora(2, " DataHoraTermino_ProcessamentoNFCeComplementar = ", mDataHoraTermino)
    Call sqlTexto(2, ", Observacao_ProcessamentoNFCeComplementar = ", pObservacao)


    gSQL = gSQL & " WHERE IdEstabelecimento_ProcessamentoNFCeComplementar = " & pIdEstabelecimento
    gSQL = gSQL & " AND DataHoraInicio_ProcessamentoNFCeComplementar = " & preparaDataHora(pDataHoraInicio)
    gSQL = gSQL & " AND PDV_ProcessamentoNFCeComplementar = " & pPDV


    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    
    If lRecordsAffected > 0 Then
        DefinirDataHoraTermino = True
    End If


    Exit Function

Err_Alterar:
End Function

Public Function DefinirDataHoraTerminoPendentes(ByVal pIdEstabelecimento As Integer, ByVal pDataHora As Date, ByVal pObservacao As String) As Boolean

Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    DefinirDataHoraTerminoPendentes = False
    gSQL = "UPDATE ProcessamentoNFCeComplementar SET "
    Call sqlDataHora(2, " DataHoraTermino_ProcessamentoNFCeComplementar = ", pDataHora)
    Call sqlTexto(2, ", Observacao_ProcessamentoNFCeComplementar = ", pObservacao)
    gSQL = gSQL & " WHERE IdEstabelecimento_ProcessamentoNFCeComplementar = " & pIdEstabelecimento
    gSQL = gSQL & " AND DataHoraTermino_ProcessamentoNFCeComplementar IS NULL "

    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    
    If lRecordsAffected > 0 Then
        DefinirDataHoraTerminoPendentes = True
    End If


    Exit Function

Err_Alterar:
End Function

Public Function DefinirDataHoraEmissao(ByVal pIdEstabelecimento As Integer, ByVal pDataHoraInicio As Date, ByVal pPDV As Integer) As Boolean

Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    DefinirDataHoraEmissao = False
    gSQL = "UPDATE ProcessamentoNFCeComplementar SET "
    Call sqlDataHora(2, " DataHoraEmissao_ProcessamentoNFCeComplementar = ", mDataHoraEmissao)


    gSQL = gSQL & " WHERE IdEstabelecimento_ProcessamentoNFCeComplementar = " & pIdEstabelecimento
    gSQL = gSQL & " AND DataHoraInicio_ProcessamentoNFCeComplementar = " & preparaDataHora(pDataHoraInicio)
    gSQL = gSQL & " AND PDV_ProcessamentoNFCeComplementar = " & pPDV


    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    
    If lRecordsAffected > 0 Then
        DefinirDataHoraEmissao = True
    End If


    Exit Function

Err_Alterar:
End Function

Public Function ExisteProcessamentoEmAndamento(ByVal pIdEstabelecimento As Integer) As Boolean
Dim xCondicao As String
Dim xDataInicial As String
Dim xDataFinal As String

On Error GoTo trata_erro

    xDataInicial = Format(Now, "dd/MM/yyyy") & " " & "00:00:00"
    xDataFinal = Format(Now, "dd/MM/yyyy") & " " & "23:59:59"
    

    ExisteProcessamentoEmAndamento = False
   
    xCondicao = xCondicao & " WHERE IdEstabelecimento_ProcessamentoNFCeComplementar = " & pIdEstabelecimento
    xCondicao = xCondicao & " AND DataHoraInicio_ProcessamentoNFCeComplementar >= " & preparaDataHora(CDate(xDataInicial))
    xCondicao = xCondicao & " AND DataHoraInicio_ProcessamentoNFCeComplementar <= " & preparaDataHora(CDate(xDataFinal))
    xCondicao = xCondicao & " AND DataHoraTermino_ProcessamentoNFCeComplementar IS NULL "

    
    gSQL = PreparaSQL(xCondicao, "")
    ExisteProcessamentoEmAndamento = Localizar(1)
    
    rsProcessa.Close
    Set rsProcessa = Nothing
    
    Exit Function

trata_erro:
    CriaLogCadastroDll ("[ExisteProcessamentoEmAndamento] - " & Err.Description)
'    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarCodigo(ByVal pIdEstabelecimento As Integer, ByVal pDataHoraInicio As Date, ByVal pPDV As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
   
    xCondicao = xCondicao & " WHERE IdEstabelecimento_ProcessamentoNFCeComplementar = " & pIdEstabelecimento
    xCondicao = xCondicao & " AND DataHoraInicio_ProcessamentoNFCeComplementar = " & preparaData(pDataHoraInicio)
    xCondicao = xCondicao & " AND PDV_ProcessamentoNFCeComplementar = " & pPDV

    
    xCondicao = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsProcessa.Close
    Set rsProcessa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function



Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    Dim xSQL As String
    
    
    xSQL = "SELECT IdEstabelecimento_ProcessamentoNFCeComplementar, DataHoraInicio_ProcessamentoNFCeComplementar, PDV_ProcessamentoNFCeComplementar, "
    xSQL = xSQL & " DataHoraEmissao_ProcessamentoNFCeComplementar, DataHoraTermino_ProcessamentoNFCeComplementar, Observacao_ProcessamentoNFCeComplementar "
    xSQL = xSQL & " FROM ProcessamentoNFCeComplementar"
    
    PreparaSQL = Trim(xSQL & " " & xCondicao & " " & xTipoOrdem)
End Function

