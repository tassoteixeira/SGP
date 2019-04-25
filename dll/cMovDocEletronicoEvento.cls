VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovDocEletronicoEvento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Private mIdEstabelecimento As Integer
    Private mDataEmissao As Date
    Private mModelo As String
    Private mSerie As String
    Private mNumero As Long
    Private mSequencia As Integer
    Private mDataHora As Date
    Private mCodigoTipoEvento As Integer
    Private mDescricao As String
    
    
    Private rsEvento As New adodb.Recordset

    Public Enum EVENTO_NFCE
        NENHUM_EVENTO = 0
        ABERTA = 100
        ABERTA_COMPLEMENTAR_AUTO = 150
        ABERTA_COMPLEMENTAR = 160
        ABANDONADA = 200
        FECHADA = 300
        PROCESSAMENTO_INICIADO = 400
        PROCESSAMENTO_CONCLUIDO_NEGADA = 500
        PROCESSAMENTO_CONCLUIDO_SEM_RETORNO = 550
        PROCESSAMENTO_CONCLUIDO_AUTORIZADA_CONTINGENCIA = 600
        PROCESSAMENTO_CONCLUIDO_AUTORIZADA = 700
        CONCILIADA_DIVERGENCIA = 800
        CONCILIADA_OK = 900
    End Enum

    Public Function DescricaoEnumEvento(ByVal pValorEnum As EVENTO_NFCE) As String

        DescricaoEnumEvento = ""
        Select Case pValorEnum

            Case EVENTO_NFCE.NENHUM_EVENTO
                DescricaoEnumEvento = "NENHUM_EVENTO"
            Case EVENTO_NFCE.ABERTA
                DescricaoEnumEvento = "ABERTA"
            Case EVENTO_NFCE.ABERTA_COMPLEMENTAR
                DescricaoEnumEvento = "ABERTA_COMPLEMENTAR"
            Case EVENTO_NFCE.ABERTA_COMPLEMENTAR_AUTO
                DescricaoEnumEvento = "ABERTA_COMPLEMENTAR_AUTO"
            Case EVENTO_NFCE.ABANDONADA
                DescricaoEnumEvento = "ABANDONADA"
            Case EVENTO_NFCE.FECHADA
                DescricaoEnumEvento = "FECHADA"
            Case EVENTO_NFCE.PROCESSAMENTO_INICIADO
                DescricaoEnumEvento = "PROCESSAMENTO_INICIADO"
            Case EVENTO_NFCE.PROCESSAMENTO_CONCLUIDO_NEGADA
                DescricaoEnumEvento = "PROCESSAMENTO_CONCLUIDO_NEGADA"
            Case EVENTO_NFCE.PROCESSAMENTO_CONCLUIDO_SEM_RETORNO
                DescricaoEnumEvento = "PROCESSAMENTO_CONCLUIDO_SEM_RETORNO"
            Case EVENTO_NFCE.PROCESSAMENTO_CONCLUIDO_AUTORIZADA_CONTINGENCIA
                DescricaoEnumEvento = "PROCESSAMENTO_CONCLUIDO_AUTORIZADA_CONTINGENCIA"
            Case EVENTO_NFCE.PROCESSAMENTO_CONCLUIDO_AUTORIZADA
                DescricaoEnumEvento = "PROCESSAMENTO_CONCLUIDO_AUTORIZADA"
            Case EVENTO_NFCE.CONCILIADA_DIVERGENCIA
                DescricaoEnumEvento = "CONCILIADA_DIVERGENCIA"
            Case EVENTO_NFCE.CONCILIADA_OK
                DescricaoEnumEvento = "CONCILIADA_OK"
        End Select

    End Function

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

Public Property Get Numero() As Long
    Numero = mNumero
End Property
Public Property Let Numero(ByVal Valor As Long)
    mNumero = Valor
End Property

Public Property Get Sequencia() As Integer
    Sequencia = mSequencia
End Property
Public Property Let Sequencia(ByVal Valor As Integer)
    mSequencia = Valor
End Property

Public Property Get DataHora() As Date
    DataHora = mDataHora
End Property
Public Property Let DataHora(ByVal Valor As Date)
    mDataHora = Valor
End Property

Public Property Get Descricao() As String
    Descricao = mDescricao
End Property
Public Property Let Descricao(ByVal Valor As String)
    mDescricao = Valor
End Property
Public Property Get CodigoTipoEvento() As Integer
    CodigoTipoEvento = mCodigoTipoEvento
End Property
Public Property Let CodigoTipoEvento(ByVal Valor As Integer)
    mCodigoTipoEvento = Valor
End Property


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False

    gSQL = "INSERT INTO MovDocumentoEletronicoEvento (IdEstabelecimento_MovDEEvento, DataEmissao_MovDEEvento, Modelo_MovDEEvento, "
    gSQL = gSQL & " Serie_MovDEEvento, Numero_MovDEEvento, Sequencia_MovDEEvento, DataHora_MovDEEvento, CodigoTipoEvento_MovDEEvento, Descricao_MovDEEvento "
    gSQL = gSQL & ") VALUES ("
    Call sqlNumero(1, mIdEstabelecimento, ", ")
    Call sqlData(1, mDataEmissao, ", ")
    Call sqlTexto(1, mModelo, ", ")
    Call sqlTexto(1, mSerie, ", ")
    Call sqlNumero(1, mNumero, ", ")
    Call sqlNumero(1, mSequencia, ", ")
    Call sqlDataHora(1, mDataHora, ", ")
    Call sqlNumero(1, mCodigoTipoEvento, ", ")
    Call sqlTexto(1, mDescricao, " )")
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    
    If lRecordsAffected > 0 Then
       Incluir = True
    End If
    Exit Function
    
Err_Incluir:
End Function

Public Function Alterar(ByVal pIdEstabelecimento As Integer, ByVal pDataEmissao As Date, ByVal pModelo As String, ByVal pSerie As String, ByVal pNumero As Long, ByVal pSequencia As Integer) As Boolean
    Dim lRecordsAffected As Long
    
    On Error GoTo Err_Alterar
    
    Alterar = False
    
    gSQL = "UPDATE MovDocumentoEletronicoEvento SET "
    Call sqlNumero(2, "IdEstabelecimento_MovDEEvento = ", mIdEstabelecimento)
    Call sqlData(2, ", DataEmissao_MovDEEvento = ", mDataEmissao)
    Call sqlTexto(2, ", Modelo_MovDEEvento = ", mModelo)
    Call sqlTexto(2, ", Serie_MovDEEvento = ", mSerie)
    Call sqlNumero(2, ", Numero_MovDEEvento = ", mNumero)
    Call sqlNumero(2, ", Sequencia_MovDEEvento = ", mSequencia)
    Call sqlDataHora(2, ", DataHora_MovDEEvento = ", mDataHora)
    Call sqlNumero(2, ", CodigoTipoEvento_MovDEEvento = ", mCodigoTipoEvento)
    Call sqlTexto(2, ", Descricao_MovDEEvento = ", mDescricao)
    
    gSQL = gSQL & " WHERE IdEstabelecimento_MovDEEvento = " & pIdEstabelecimento
    gSQL = gSQL & " AND DataEmissao_MovDEEvento = " & preparaData(pDataEmissao)
    gSQL = gSQL & " AND Modelo_MovDEEvento = " & preparaTexto(pModelo)
    gSQL = gSQL & " AND Serie_MovDEEvento = " & preparaTexto(pSerie)
    gSQL = gSQL & " AND Numero_MovDEEvento = " & pNumero
    gSQL = gSQL & " AND Sequencia_MovDEEvento = " & pSequencia
    
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
    Set rsEvento = New adodb.Recordset
    rsEvento.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsEvento.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsEvento.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()

    IdEstabelecimento = rsEvento("IdEstabelecimento_MovDEEvento").Value
    DataEmissao = rsEvento("DataEmissao_MovDEEvento").Value
    Modelo = rsEvento("Modelo_MovDEEvento").Value
    Serie = rsEvento("Serie_MovDEEvento").Value
    Numero = rsEvento("Numero_MovDEEvento").Value
    Sequencia = rsEvento("Sequencia_MovDEEvento").Value
    DataHora = rsEvento("DataHora_MovDEEvento").Value
    CodigoTipoEvento = rsEvento("CodigoTipoEvento_MovDEEvento").Value
    Descricao = rsEvento("Descricao_MovDEEvento").Value
End Sub


Public Function LocalizarCodigo(ByVal pIdEstabelecimento As Integer, ByVal pDataEmissao As Date, ByVal pModelo As String, ByVal pSerie As String, ByVal pNumero As String, ByVal pSequencia As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
   
    xCondicao = xCondicao & " WHERE IdEstabelecimento_MovDEEvento = " & pIdEstabelecimento
    xCondicao = xCondicao & " AND DataEmissao_MovDEEvento = " & preparaData(pDataEmissao)
    xCondicao = xCondicao & " AND Modelo_MovDEEvento = " & preparaTexto(pModelo)
    xCondicao = xCondicao & " AND Serie_MovDEEvento = " & preparaTexto(pSerie)
    xCondicao = xCondicao & " AND Numero_MovDEEvento = " & pNumero
    xCondicao = xCondicao & " AND Sequencia_MovDEEvento = " & pSequencia

    
    xCondicao = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsEvento.Close
    Set rsEvento = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

Public Function ProximaSequencia(ByVal pIdEstabelecimento As Integer, ByVal pDataEmissao As Date, ByVal pModelo As String, ByVal pSerie As String, ByVal pNumero As String) As Integer
Dim xCondicao As String
Dim xLocalizouProximaSequencia As Boolean
On Error GoTo trata_erro


    ProximaSequencia = 1
    xLocalizouProximaSequencia = False
    gSQL = ""
    xCondicao = ""
    
    xCondicao = xCondicao & " WHERE IdEstabelecimento_MovDEEvento = " & pIdEstabelecimento
    xCondicao = xCondicao & " AND DataEmissao_MovDEEvento = " & preparaData(pDataEmissao)
    xCondicao = xCondicao & " AND Modelo_MovDEEvento = " & preparaTexto(pModelo)
    xCondicao = xCondicao & " AND Serie_MovDEEvento = " & preparaTexto(pSerie)
    xCondicao = xCondicao & " AND Numero_MovDEEvento = " & pNumero

    
    gSQL = PreparaSQL(xCondicao, " ORDER BY Sequencia_MovDEEvento DESC")
    xLocalizouProximaSequencia = Localizar(1)
    If xLocalizouProximaSequencia Then
        ProximaSequencia = rsEvento("Sequencia_MovDEEvento").Value + 1
    End If
    
    rsEvento.Close
    Set rsEvento = Nothing
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
    xSQL = "SELECT IdEstabelecimento_MovDEEvento, DataEmissao_MovDEEvento, Modelo_MovDEEvento, "
    xSQL = xSQL & " Serie_MovDEEvento, Numero_MovDEEvento, Sequencia_MovDEEvento, DataHora_MovDEEvento, CodigoTipoEvento_MovDEEvento, Descricao_MovDEEvento "
    xSQL = xSQL & " FROM MovDocumentoEletronicoEvento"
    
    PreparaSQL = Trim(xSQL & " " & xCondicao & " " & xTipoOrdem)
End Function

