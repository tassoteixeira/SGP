VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLiberacaoDigitacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mDataInicial As Date
Private mDataFinal As Date
Private mPeriodoInicial As String
Private mPeriodoFinal As String
Private mTipoMovimento As Integer

Private rsLiberacaoDigitacao As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get DataInicial() As Date
    DataInicial = mDataInicial
End Property
Public Property Let DataInicial(ByVal Valor As Date)
    mDataInicial = Valor
End Property


Public Property Get DataFinal() As Date
    DataFinal = mDataFinal
End Property
Public Property Let DataFinal(ByVal Valor As Date)
    mDataFinal = Valor
End Property


Public Property Get PeriodoInicial() As String
    PeriodoInicial = mPeriodoInicial
End Property
Public Property Let PeriodoInicial(ByVal Valor As String)
    mPeriodoInicial = Valor
End Property


Public Property Get PeriodoFinal() As String
    PeriodoFinal = mPeriodoFinal
End Property
Public Property Let PeriodoFinal(ByVal Valor As String)
    mPeriodoFinal = Valor
End Property


Public Property Get TipoMovimento() As Integer
    TipoMovimento = mTipoMovimento
End Property
Public Property Let TipoMovimento(ByVal Valor As Integer)
    mTipoMovimento = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pTipoMovimento As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Tipo do Movimento] = " & pTipoMovimento
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsLiberacaoDigitacao.Close
    Set rsLiberacaoDigitacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Liberacao_Digitacao ( Empresa, [Data Inicial], [Data Final], [Periodo Inicial], "
    gSQL = gSQL & "[Periodo Final], [Tipo do Movimento] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mDataInicial, ", ")
    Call sqlData(1, mDataFinal, ", ")
    Call sqlTexto(1, mPeriodoInicial, ", ")
    Call sqlTexto(1, mPeriodoFinal, ", ")
    Call sqlNumero(1, mTipoMovimento, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pTipoMovimento As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Liberacao_Digitacao SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", [Data Inicial] = ", mDataInicial)
    Call sqlData(2, ", [Data Final] = ", mDataFinal)
    Call sqlTexto(2, ", [Periodo Inicial] = ", mPeriodoInicial)
    Call sqlTexto(2, ", [Periodo Final] = ", mPeriodoFinal)
    Call sqlNumero(2, ", [Tipo do Movimento] = ", mTipoMovimento)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Tipo do Movimento] = " & pTipoMovimento
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
        Call CriaLogCadastroDll2("cLiberacaoDigitacao:Alterar - ** Investigação Especial - Registro alterado p/:", "Investigação Especial", gSQL)
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pTipoMovimento As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Liberacao_Digitacao"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Tipo do Movimento] = " & pTipoMovimento
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Data Inicial], [Data Final], [Periodo Inicial], [Periodo Final], [Tipo do Movimento] FROM Liberacao_Digitacao" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsLiberacaoDigitacao = New adodb.Recordset
    rsLiberacaoDigitacao.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsLiberacaoDigitacao.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsLiberacaoDigitacao.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsLiberacaoDigitacao("Empresa").Value
    DataInicial = rsLiberacaoDigitacao("Data Inicial").Value
    DataFinal = rsLiberacaoDigitacao("Data Final").Value
    PeriodoInicial = rsLiberacaoDigitacao("Periodo Inicial").Value
    PeriodoFinal = rsLiberacaoDigitacao("Periodo Final").Value
    TipoMovimento = rsLiberacaoDigitacao("Tipo do Movimento").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

