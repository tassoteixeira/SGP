VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFechamentoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mData As Date
Private mPeriodo As Integer
Private mCupomInicial As Long
Private mCupomFinal As Long
Private mHoraInicial As Date
Private mHoraFinal As Date

Private rsFechamentoCaixa As New adodb.Recordset


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


Public Property Get Periodo() As Integer
    Periodo = mPeriodo
End Property
Public Property Let Periodo(ByVal Valor As Integer)
    mPeriodo = Valor
End Property


Public Property Get CupomInicial() As Long
    CupomInicial = mCupomInicial
End Property
Public Property Let CupomInicial(ByVal Valor As Long)
    mCupomInicial = Valor
End Property


Public Property Get CupomFinal() As Long
    CupomFinal = mCupomFinal
End Property
Public Property Let CupomFinal(ByVal Valor As Long)
    mCupomFinal = Valor
End Property


Public Property Get HoraInicial() As Date
    HoraInicial = mHoraInicial
End Property
Public Property Let HoraInicial(ByVal Valor As Date)
    mHoraInicial = Valor
End Property


Public Property Get HoraFinal() As Date
    HoraFinal = mHoraFinal
End Property
Public Property Let HoraFinal(ByVal Valor As Date)
    mHoraFinal = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nPeriodo As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(nData)
    xCondicao = xCondicao & " AND Periodo = " & nPeriodo
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsFechamentoCaixa.Close
    Set rsFechamentoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnteriorA(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pPeriodo As Integer) As Boolean
Dim xCondicao As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarAnteriorA = False
    xRegIgual = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data <= " & preparaData(pData)
    xCondicao = xCondicao & " AND Periodo <= " & pPeriodo
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, Periodo DESC")
    If Localizar(55) Then
        If rsFechamentoCaixa.RecordCount > 1 Then
            Do Until rsFechamentoCaixa.EOF
                If xRegIgual = False Then
                    If rsFechamentoCaixa("Data").Value = mData And rsFechamentoCaixa("Periodo").Value = mPeriodo Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnteriorA = True
                    Exit Do
                End If
                rsFechamentoCaixa.MoveNext
            Loop
        End If
    End If
    If LocalizarAnteriorA = True Then
        AtribuiValor
    End If
    rsFechamentoCaixa.Close
    Set rsFechamentoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, Periodo DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsFechamentoCaixa.Close
    Set rsFechamentoCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Fechamento_Caixa ( Empresa, Data, Periodo, [Cupom Inicial], [Cupom Final], "
    gSQL = gSQL & "[Hora Inicial], [Hora Final] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlNumero(1, mPeriodo, ", ")
    Call sqlNumero(1, mCupomInicial, ", ")
    Call sqlNumero(1, mCupomFinal, ", ")
    Call sqlData(1, mHoraInicial, ", ")
    Call sqlData(1, mHoraFinal, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nPeriodo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Fechamento_Caixa SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlNumero(2, ", Periodo = ", mPeriodo)
    Call sqlNumero(2, ", [Cupom Inicial] = ", mCupomInicial)
    Call sqlNumero(2, ", [Cupom Final] = ", mCupomFinal)
    Call sqlData(2, ", [Hora Inicial] = ", mHoraInicial)
    Call sqlData(2, ", [Hora Final] = ", mHoraFinal)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND Periodo = " & nPeriodo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nPeriodo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Fechamento_Caixa"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND Periodo = " & nPeriodo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, Periodo, [Cupom Inicial], [Cupom Final], [Hora Inicial], [Hora Final] FROM Fechamento_Caixa" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsFechamentoCaixa = New adodb.Recordset
    rsFechamentoCaixa.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsFechamentoCaixa.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsFechamentoCaixa.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsFechamentoCaixa("Empresa").Value
    Data = rsFechamentoCaixa("Data").Value
    Periodo = rsFechamentoCaixa("Periodo").Value
    CupomInicial = rsFechamentoCaixa("Cupom Inicial").Value
    CupomFinal = rsFechamentoCaixa("Cupom Final").Value
    HoraInicial = rsFechamentoCaixa("Hora Inicial").Value
    HoraFinal = rsFechamentoCaixa("Hora Final").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

