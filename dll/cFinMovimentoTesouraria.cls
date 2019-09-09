VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFinMovimentoTesouraria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Private mEmpresa As Integer
    Private mData As Date
    Private mNumeroMovimento As Long
    Private mCodigoTipoMovimento As Integer
    Private mHistorico As String
    Private mValorEntrada As Currency
    Private mValorSaida As Currency
    Private mCodigoContaTesouraria As Integer
    Private mRegistroContaPagar As Integer
    
    Private rsMovimentoTesouraria As New adodb.Recordset

    
'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property
Public Property Get Data() As Date
    Data = mData
End Property
Public Property Let Data(ByVal Valor As Date)
    mData = Valor
End Property
Public Property Get NumeroMovimento() As Long
    NumeroMovimento = mNumeroMovimento
End Property
Public Property Let NumeroMovimento(ByVal Valor As Long)
    mNumeroMovimento = Valor
End Property
Public Property Get CodigoTipoMovimento() As Integer
    CodigoTipoMovimento = mCodigoTipoMovimento
End Property
Public Property Let CodigoTipoMovimento(ByVal Valor As Integer)
    mCodigoTipoMovimento = Valor
End Property
Public Property Get Historico() As String
    Historico = mHistorico
End Property
Public Property Let Historico(ByVal Valor As String)
    mHistorico = Valor
End Property
Public Property Get ValorEntrada() As Currency
    ValorEntrada = mValorEntrada
End Property
Public Property Let ValorEntrada(ByVal Valor As Currency)
    mValorEntrada = Valor
End Property
Public Property Get ValorSaida() As Currency
    ValorSaida = mValorSaida
End Property
Public Property Let ValorSaida(ByVal Valor As Currency)
    mValorSaida = Valor
End Property
Public Property Get CodigoContaTesouraria() As Integer
    CodigoContaTesouraria = mCodigoContaTesouraria
End Property
Public Property Let CodigoContaTesouraria(ByVal Valor As Integer)
    mCodigoContaTesouraria = Valor
End Property
Public Property Get RegistroContaPagar() As Integer
    RegistroContaPagar = mRegistroContaPagar
End Property
Public Property Let RegistroContaPagar(ByVal Valor As Integer)
    mRegistroContaPagar = Valor
End Property



Public Function Alterar(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumeroMovimento As Long) As Boolean
    Dim lRecordsAffected As Long

    On Error GoTo Err_Alterar
        Alterar = False
            gSQL = "UPDATE FinMovimentoTesouraria SET "
            Call sqlNumero(2, "Empresa = ", mEmpresa)
            Call sqlData(2, ", Data = ", mData)
            Call sqlNumero(2, ", [Numero do Movimento] = ", mNumeroMovimento)
            Call sqlNumero(2, ", [Codigo do Tipo de Movimento] = ", mCodigoTipoMovimento)
            Call sqlTexto(2, ", Historico = ", mHistorico)
            Call sqlValor(2, ", [Valor de Entrada] = ", mValorEntrada)
            Call sqlValor(2, ", [Valor de Saida] = ", mValorSaida)
            Call sqlNumero(2, ", [Codigo da Conta de Tesouraria] = ", mCodigoContaTesouraria)
            Call sqlNumero(2, ", RegistroContaPagar = ", mRegistroContaPagar)
            gSQL = gSQL & " WHERE Empresa = " & pEmpresa
            gSQL = gSQL & "   AND Data = " & preparaData(pData)
            gSQL = gSQL & "   AND [Numero do Movimento] = " & pNumeroMovimento
    
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
            
        If lRecordsAffected > 0 Then
            Alterar = True
        Else
            Call CriaLogCadastroDll2("FinMovimentoTesouraria:Alterar - Erro ao alterar registro.", Err.Description, gSQL)
        End If
        Exit Function

Err_Alterar:
    Call CriaLogCadastroDll2("FinMovimentoTesouraria:Alterar - Erro não identificado.", Err.Description, gSQL)
    Exit Function
           
End Function

Public Function Excluir(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumeroMovimento As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir
        Excluir = False
            
            gSQL = "DELETE FROM FinMovimentoTesouraria"
            gSQL = gSQL & " WHERE Empresa = " & pEmpresa
            gSQL = gSQL & "   AND Data = " & preparaData(pData)
            gSQL = gSQL & "   AND [Numero do Movimento] = " & pNumeroMovimento
            
            gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        If lRecordsAffected > 0 Then
            Excluir = True
        End If
Exit Function

Err_Excluir:
    Call CriaLogCadastroDll2("FinMovimentoTesouraria:Excluir - Erro não identificado.", Err.Description, gSQL)
    Exit Function

End Function
Public Function Incluir() As Boolean
    Dim lRecordsAffected As Long

    On Error GoTo Err_Incluir
        Incluir = False
            gSQL = "INSERT INTO FinMovimentoTesouraria ( Empresa, Data, [Numero do Movimento], [Codigo do Tipo de Movimento], "
            gSQL = gSQL & "Historico, [Valor de Entrada], [Valor de Saida], "
            gSQL = gSQL & "[Codigo da Conta de Tesouraria], RegistroContaPagar ) VALUES ( "
            Call sqlNumero(1, mEmpresa, ", ")
            Call sqlData(1, mData, ", ")
            Call sqlNumero(1, mNumeroMovimento, ", ")
            Call sqlNumero(1, mCodigoTipoMovimento, ", ")
            Call sqlTexto(1, mHistorico, ", ")
            Call sqlValor(1, mValorEntrada, ", ")
            Call sqlValor(1, mValorSaida, ", ")
            Call sqlNumero(1, CodigoContaTesouraria, ", ")
            Call sqlNumero(1, RegistroContaPagar, " )")
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
    Call CriaLogCadastroDll2("FinMovimentoTesouraria:Incluir - Erro não identificado.", Err.Description, gSQL)
    Exit Function

End Function

'Funções / Procedures internas
    Private Function PreparaSQL(ByVal pCondicao As String, ByVal pTipoOrdem As String) As String
        PreparaSQL = Trim("SELECT Empresa, Data, [Numero do Movimento], [Codigo do Tipo de Movimento], Historico, [Valor de Entrada], [Valor de Saida], [Codigo da Conta de Tesouraria], RegistroContaPagar FROM FinMovimentoTesouraria" & " " & pCondicao & " " & pTipoOrdem)
    End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovimentoTesouraria = New adodb.Recordset
    rsMovimentoTesouraria.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovimentoTesouraria.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovimentoTesouraria.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    
        Empresa = rsMovimentoTesouraria("Empresa").Value
        Data = rsMovimentoTesouraria("Data").Value
        NumeroMovimento = rsMovimentoTesouraria("Numero do Movimento").Value
        CodigoTipoMovimento = rsMovimentoTesouraria("Codigo do Tipo de Movimento").Value
        Historico = rsMovimentoTesouraria("Historico").Value
        ValorEntrada = rsMovimentoTesouraria("Valor de Entrada").Value
        ValorSaida = rsMovimentoTesouraria("Valor de Saida").Value
        CodigoContaTesouraria = rsMovimentoTesouraria("Codigo da Conta de Tesouraria").Value
        RegistroContaPagar = rsMovimentoTesouraria("RegistroContaPagar").Value
    
End Sub

Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
        Call AtribuiValor
    End If
    rsMovimentoTesouraria.Close
    Set rsMovimentoTesouraria = Nothing

End Function


Public Function ProximoCodigo(ByVal pEmpresa As Integer, ByVal pData As Date) As Long
    Dim xCondicao As String
    On Error GoTo trata_erro
        ProximoCodigo = 1
        xCondicao = "SELECT [Numero do Movimento]"
        xCondicao = xCondicao & "  FROM FinMovimentoTesouraria"
        xCondicao = xCondicao & " WHERE Empresa = " & pEmpresa
        xCondicao = xCondicao & "   AND Data = " & preparaData(pData)
        xCondicao = xCondicao & " ORDER BY [Numero do Movimento] DESC"
        rsMovimentoTesouraria.Open xCondicao, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovimentoTesouraria.EOF Then
        ProximoCodigo = rsMovimentoTesouraria("Numero do Movimento").Value + 1
    End If
    
    rsMovimentoTesouraria.Close
    
    Set rsMovimentoTesouraria = Nothing
    
    Exit Function
trata_erro:
    MsgBox Err.Number & " - " & Err.Description
    Exit Function
End Function
Public Function LocalizarRegistroContaAPagar(ByVal pEmpresa As Integer, ByVal pRegistroContaAPagar As Integer) As Boolean
    Dim xCondicao As String
    On Error GoTo trata_erro
        LocalizarRegistroContaAPagar = False
        xCondicao = " WHERE Empresa  = " & pEmpresa
        xCondicao = xCondicao & " AND RegistroContaPagar  = " & pRegistroContaAPagar
       
        gSQL = PreparaSQL(xCondicao, "")
        
        
        LocalizarRegistroContaAPagar = Localizar(1)

    
    Exit Function
trata_erro:
    MsgBox Err.Number & " - " & Err.Description
    Exit Function
End Function
    

