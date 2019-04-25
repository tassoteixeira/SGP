VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMedicaoCombustivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mNomeTabela As String
Private mEmpresa As Integer
Private mData As Date
Private mNumeroTanque As Integer
Private mTipoCombustivel As String
Private mQuantidade As Currency
Private mObservacao1 As String
Private mObservacao2 As String
Private mObservacao3 As String
Private mDescontoDiaAnterior As Currency
Private mMedidaRegua As Currency

Private rsMedicaoCombustivel As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get NomeTabela() As String
    NomeTabela = mNomeTabela
End Property
Public Property Let NomeTabela(ByVal Valor As String)
    mNomeTabela = Valor
End Property


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


Public Property Get NumeroTanque() As Integer
    NumeroTanque = mNumeroTanque
End Property
Public Property Let NumeroTanque(ByVal Valor As Integer)
    mNumeroTanque = Valor
End Property


Public Property Get TipoCombustivel() As String
    TipoCombustivel = mTipoCombustivel
End Property
Public Property Let TipoCombustivel(ByVal Valor As String)
    mTipoCombustivel = Valor
End Property


Public Property Get Quantidade() As Currency
    Quantidade = mQuantidade
End Property
Public Property Let Quantidade(ByVal Valor As Currency)
    mQuantidade = Valor
End Property


Public Property Get Observacao1() As String
    Observacao1 = mObservacao1
End Property
Public Property Let Observacao1(ByVal Valor As String)
    mObservacao1 = Valor
End Property


Public Property Get Observacao2() As String
    Observacao2 = mObservacao2
End Property
Public Property Let Observacao2(ByVal Valor As String)
    mObservacao2 = Valor
End Property


Public Property Get Observacao3() As String
    Observacao3 = mObservacao3
End Property
Public Property Let Observacao3(ByVal Valor As String)
    mObservacao3 = Valor
End Property


Public Property Get DescontoDiaAnterior() As Currency
    DescontoDiaAnterior = mDescontoDiaAnterior
End Property
Public Property Let DescontoDiaAnterior(ByVal Valor As Currency)
    mDescontoDiaAnterior = Valor
End Property


Public Property Get MedidaRegua() As Currency
    MedidaRegua = mMedidaRegua
End Property
Public Property Let MedidaRegua(ByVal Valor As Currency)
    mMedidaRegua = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumeroTanque As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND [Numero do Tanque] = " & pNumeroTanque
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsMedicaoCombustivel.Close
    Set rsMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarDataCombustivel(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTipoCombustivel As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarDataCombustivel = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Tanque]")
    If Localizar(1) Then
        LocalizarDataCombustivel = True
        AtribuiValor
    End If
    rsMedicaoCombustivel.Close
    Set rsMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior(ByVal nEmpresa As Integer, ByVal nData As Date) As Boolean
Dim xCondicao As String
Dim xOrdem As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Data < " & preparaData(nData)
    xOrdem = "ORDER BY Empresa DESC, Data DESC, [Numero do Tanque] DESC"
    gSQL = PreparaSQL(xCondicao, xOrdem)
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsMedicaoCombustivel.Close
    Set rsMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Data ASC, [Numero do Tanque] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsMedicaoCombustivel.Close
    Set rsMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Localiza o Primeiro Tanque Com Medida
'de um determinado Tipo de Combustível
'em uma data específica
Public Function LocalizarPrimeiroTanqueComb(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTipoCombustivel As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiroTanqueComb = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Tanque]")
    LocalizarPrimeiroTanqueComb = Localizar(1)
    If LocalizarPrimeiroTanqueComb = True Then
        AtribuiValor
    End If
    rsMedicaoCombustivel.Close
    Set rsMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Localiza o Próximo Tanque Com Medida
'de um determinado Tipo de Combustível
'em uma data específica
Public Function LocalizarProximoTanqueComb(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTipoCombustivel As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximoTanqueComb = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    xCondicao = xCondicao & " AND [Numero do Tanque] > " & mNumeroTanque
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Tanque]")
    LocalizarProximoTanqueComb = Localizar(1)
    If LocalizarProximoTanqueComb = True Then
        AtribuiValor
    End If
    rsMedicaoCombustivel.Close
    Set rsMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo(ByVal nEmpresa As Integer, ByVal nData As Date) As Boolean
Dim xCondicao As String
Dim xOrdem As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Data > " & preparaData(nData)
    xOrdem = "ORDER BY Empresa ASC, Data ASC, [Numero do Tanque] ASC"
    gSQL = PreparaSQL(xCondicao, xOrdem)
    LocalizarProximo = Localizar(1)
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMedicaoCombustivel.Close
    Set rsMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, [Numero do Tanque] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsMedicaoCombustivel.Close
    Set rsMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO " & mNomeTabela & " ( Empresa, Data, [Numero do Tanque], [Tipo de Combustivel], "
    gSQL = gSQL & "Quantidade, [Observacao 1], [Observacao 2], [Observacao 3], [Desconto Dia Anterior], "
    gSQL = gSQL & "[Medida da Regua] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlNumero(1, mNumeroTanque, ", ")
    Call sqlTexto(1, mTipoCombustivel, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlTexto(1, mObservacao1, ", ")
    Call sqlTexto(1, mObservacao2, ", ")
    Call sqlTexto(1, mObservacao3, ", ")
    Call sqlValor(1, mDescontoDiaAnterior, ", ")
    Call sqlValor(1, mMedidaRegua, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumeroTanque As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE " & mNomeTabela & " SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlNumero(2, ", [Numero do Tanque] = ", mNumeroTanque)
    Call sqlTexto(2, ", [Tipo de Combustivel] = ", mTipoCombustivel)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlTexto(2, ", [Observacao 1] = ", mObservacao1)
    Call sqlTexto(2, ", [Observacao 2] = ", mObservacao2)
    Call sqlTexto(2, ", [Observacao 3] = ", mObservacao3)
    Call sqlValor(2, ", [Desconto Dia Anterior] = ", mDescontoDiaAnterior)
    Call sqlValor(2, ", [Medida da Regua] = ", mMedidaRegua)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Numero do Tanque] = " & pNumeroTanque
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumeroTanque As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Numero do Tanque] = " & pNumeroTanque
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


Public Function ExcluirRegistros(ByVal nEmpresa As Integer, ByVal nData As Date) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    ExcluirRegistros = False
    gSQL = "DELETE FROM " & mNomeTabela
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        ExcluirRegistros = True
    End If
    Exit Function

Err_Excluir:
End Function


Public Function TransfereDadosLMC(ByVal pEmpresa As Integer, ByVal pSomenteVerifica As Boolean) As Boolean
    Dim xData As Date
    Dim rsAuxiliar As New adodb.Recordset

On Error GoTo trata_erro

    TransfereDadosLMC = False
    
    'Busca Ultima Data do Movimento LMC
    xData = UltimaData(pEmpresa)
    If xData = "00:00:00" Then
        Exit Function
    End If
    
    'Busca Movimento de Entrada para Transferencia
    xData = xData + 1
    gSQL = ""
    gSQL = gSQL & "   SELECT Empresa, Data, [Numero do Tanque], [Tipo de Combustivel],"
    gSQL = gSQL & "          Quantidade, [Observacao 1], [Observacao 2], [Observacao 3],"
    gSQL = gSQL & "          [Desconto Dia Anterior], [Medida da Regua]"
    gSQL = gSQL & "     FROM MedicaoCombustivel"
    gSQL = gSQL & "    WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "      AND Data >= " & preparaData(xData)
    gSQL = gSQL & " ORDER BY Data, [Numero do Tanque]"
    Set rsAuxiliar = New adodb.Recordset
    rsAuxiliar.CursorLocation = adUseClient
    rsAuxiliar.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    
    'Grava Entradas
    If Not rsAuxiliar.EOF Then
        If pSomenteVerifica Then
            TransfereDadosLMC = True
            Exit Function
        End If
        Do Until rsAuxiliar.EOF
            Empresa = rsAuxiliar("Empresa").Value
            Data = rsAuxiliar("Data").Value
            NumeroTanque = rsAuxiliar("Numero do Tanque").Value
            TipoCombustivel = rsAuxiliar("Tipo de Combustivel").Value
            Quantidade = rsAuxiliar("Quantidade").Value
            Observacao1 = rsAuxiliar("Observacao 1").Value
            Observacao2 = rsAuxiliar("Observacao 2").Value
            Observacao3 = rsAuxiliar("Observacao 3").Value
            DescontoDiaAnterior = rsAuxiliar("Desconto Dia Anterior").Value
            MedidaRegua = rsAuxiliar("Medida da Regua").Value
            If Not Incluir Then
                MsgBox "Registro de Medição de Combustível não foi gravado!", vbInformation, "Erro de Integridade"
            End If
            rsAuxiliar.MoveNext
        Loop
        TransfereDadosLMC = True
    End If
    
    rsAuxiliar.Close
    Set rsAuxiliar = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function UltimaData(ByVal pEmpresa As Integer) As Date

On Error GoTo trata_erro

    UltimaData = "01/01/1900"
    gSQL = "SELECT Data FROM MedicaoCombustivelLMC WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " ORDER BY Data DESC, [Numero do Tanque] DESC"
    If Localizar(1) Then
        UltimaData = rsMedicaoCombustivel("Data").Value
    End If
    rsMedicaoCombustivel.Close
    Set rsMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, [Numero do Tanque], [Tipo de Combustivel], Quantidade, [Observacao 1], [Observacao 2], [Observacao 3], [Desconto Dia Anterior], [Medida da Regua] FROM " & mNomeTabela & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMedicaoCombustivel = New adodb.Recordset
    rsMedicaoCombustivel.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMedicaoCombustivel.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMedicaoCombustivel.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMedicaoCombustivel("Empresa").Value
    Data = rsMedicaoCombustivel("Data").Value
    NumeroTanque = rsMedicaoCombustivel("Numero do Tanque").Value
    TipoCombustivel = rsMedicaoCombustivel("Tipo de Combustivel").Value
    Quantidade = rsMedicaoCombustivel("Quantidade").Value
    Observacao1 = rsMedicaoCombustivel("Observacao 1").Value
    Observacao2 = rsMedicaoCombustivel("Observacao 2").Value
    Observacao3 = rsMedicaoCombustivel("Observacao 3").Value
    DescontoDiaAnterior = rsMedicaoCombustivel("Desconto Dia Anterior").Value
    MedidaRegua = rsMedicaoCombustivel("Medida da Regua").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function


Public Function TotalDescontoCombustivel(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pTipoCombustivel As String) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    TotalDescontoCombustivel = 0
    xSQL = ""
    xSQL = xSQL & " SELECT Sum([Desconto Dia Anterior]) AS TotalDesconto"
    xSQL = xSQL & "   FROM " & mNomeTabela
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data >= " & preparaData(pDataInicial)
    xSQL = xSQL & "    AND Data <= " & preparaData(pDataFinal)
    If pTipoCombustivel <> "" Then
        xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    End If
    Set rsMedicaoCombustivel = New adodb.Recordset
    rsMedicaoCombustivel.CursorLocation = adUseClient
    rsMedicaoCombustivel.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMedicaoCombustivel.EOF Then
        If Not IsNull(rsMedicaoCombustivel("TotalDesconto").Value) Then
            TotalDescontoCombustivel = rsMedicaoCombustivel("TotalDesconto").Value
        End If
    End If
    rsMedicaoCombustivel.Close
    Set rsMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function TotalMedidaCombustivel(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTipoCombustivel As String, ByVal pTanque As Integer) As Currency
Dim xSQL As String

On Error GoTo trata_erro

    TotalMedidaCombustivel = 0
    xSQL = ""
    xSQL = xSQL & " SELECT SUM(Quantidade) AS TotalQuantidade"
    xSQL = xSQL & "   FROM " & mNomeTabela
    xSQL = xSQL & "  WHERE Empresa = " & pEmpresa
    xSQL = xSQL & "    AND Data = " & preparaData(pData)
    xSQL = xSQL & "    AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    If pTanque > 0 Then
        xSQL = xSQL & "    AND [Numero do Tanque] = " & Str(pTanque)
    End If
    
    Set rsMedicaoCombustivel = New adodb.Recordset
    rsMedicaoCombustivel.CursorLocation = adUseClient
    rsMedicaoCombustivel.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMedicaoCombustivel.EOF Then
        If Not IsNull(rsMedicaoCombustivel("TotalQuantidade").Value) Then
            TotalMedidaCombustivel = rsMedicaoCombustivel("TotalQuantidade").Value
        End If
    End If
    rsMedicaoCombustivel.Close
    Set rsMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function



