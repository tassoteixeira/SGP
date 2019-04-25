VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovAdvertenciaSuspencao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mData As Date
Private mCodigoFuncionario As Integer
Private mAdvertenciaouSuspencao As String
Private mDia As Integer
Private mMotivo1 As String
Private mMotivo2 As String
Private mMotivo3 As String
Private mMotivo4 As String

Private rsMovAdvertenciaSuspencao As New adodb.Recordset


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


Public Property Get CodigoFuncionario() As Integer
    CodigoFuncionario = mCodigoFuncionario
End Property
Public Property Let CodigoFuncionario(ByVal Valor As Integer)
    mCodigoFuncionario = Valor
End Property


Public Property Get AdvertenciaouSuspencao() As String
    AdvertenciaouSuspencao = mAdvertenciaouSuspencao
End Property
Public Property Let AdvertenciaouSuspencao(ByVal Valor As String)
    mAdvertenciaouSuspencao = Valor
End Property


Public Property Get Dia() As Integer
    Dia = mDia
End Property
Public Property Let Dia(ByVal Valor As Integer)
    mDia = Valor
End Property


Public Property Get Motivo1() As String
    Motivo1 = mMotivo1
End Property
Public Property Let Motivo1(ByVal Valor As String)
    mMotivo1 = Valor
End Property


Public Property Get Motivo2() As String
    Motivo2 = mMotivo2
End Property
Public Property Let Motivo2(ByVal Valor As String)
    mMotivo2 = Valor
End Property


Public Property Get Motivo3() As String
    Motivo3 = mMotivo3
End Property
Public Property Let Motivo3(ByVal Valor As String)
    mMotivo3 = Valor
End Property


Public Property Get Motivo4() As String
    Motivo4 = mMotivo4
End Property
Public Property Let Motivo4(ByVal Valor As String)
    mMotivo4 = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nCodigoFuncionario As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(nData)
    xCondicao = xCondicao & " AND [Codigo do Funcionario] = " & nCodigoFuncionario
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsMovAdvertenciaSuspencao.Close
    Set rsMovAdvertenciaSuspencao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data <= " & preparaData(mData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, [Codigo do Funcionario] DESC")
    If Localizar(20) Then
        If rsMovAdvertenciaSuspencao.RecordCount > 1 Then
            Do Until rsMovAdvertenciaSuspencao.EOF
                If xRegIgual = False Then
                    If rsMovAdvertenciaSuspencao("Data").Value = mData And rsMovAdvertenciaSuspencao("Codigo do Funcionario").Value = mCodigoFuncionario Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsMovAdvertenciaSuspencao.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsMovAdvertenciaSuspencao.Close
    Set rsMovAdvertenciaSuspencao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal nNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL(" WHERE Nome = " & preparaTexto(nNome), "")
    LocalizarNome = Localizar(1)
    If LocalizarNome Then
        AtribuiValor
    End If
    rsMovAdvertenciaSuspencao.Close
    Set rsMovAdvertenciaSuspencao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, [Codigo do Funcionario] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsMovAdvertenciaSuspencao.Close
    Set rsMovAdvertenciaSuspencao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data >= " & preparaData(mData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, [Codigo do Funcionario] ASC")
    If Localizar(20) Then
        If rsMovAdvertenciaSuspencao.RecordCount > 1 Then
            Do Until rsMovAdvertenciaSuspencao.EOF
                If xRegIgual = False Then
                    If rsMovAdvertenciaSuspencao("Data").Value = mData And rsMovAdvertenciaSuspencao("Codigo do Funcionario").Value = mCodigoFuncionario Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsMovAdvertenciaSuspencao.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMovAdvertenciaSuspencao.Close
    Set rsMovAdvertenciaSuspencao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, [Codigo do Funcionario] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsMovAdvertenciaSuspencao.Close
    Set rsMovAdvertenciaSuspencao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Movimento_Advertencia_Suspencao ( Empresa, Data, [Codigo do Funcionario], [Advertencia ou Suspencao], Dia, "
    gSQL = gSQL & "Motivo_1, Motivo_2, Motivo_3, Motivo_4 ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlNumero(1, mCodigoFuncionario, ", ")
    Call sqlTexto(1, mAdvertenciaouSuspencao, ", ")
    Call sqlNumero(1, mDia, ", ")
    Call sqlTexto(1, mMotivo1, ", ")
    Call sqlTexto(1, mMotivo2, ", ")
    Call sqlTexto(1, mMotivo3, ", ")
    Call sqlTexto(1, mMotivo4, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nCodigoFuncionario As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Movimento_Advertencia_Suspencao SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlNumero(2, ", [Codigo do Funcionario] = ", mCodigoFuncionario)
    Call sqlTexto(2, ", [Advertencia ou Suspencao] = ", mAdvertenciaouSuspencao)
    Call sqlNumero(2, ", Dia = ", mDia)
    Call sqlTexto(2, ", Motivo_1 = ", mMotivo1)
    Call sqlTexto(2, ", Motivo_2 = ", mMotivo2)
    Call sqlTexto(2, ", Motivo_3 = ", mMotivo3)
    Call sqlTexto(2, ", Motivo_4 = ", mMotivo4)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND [Codigo do Funcionario] = " & nCodigoFuncionario
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nCodigoFuncionario As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Movimento_Advertencia_Suspencao"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND [Codigo do Funcionario] = " & nCodigoFuncionario
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, [Codigo do Funcionario], [Advertencia ou Suspencao], Dia, Motivo_1, Motivo_2, Motivo_3, Motivo_4 FROM Movimento_Advertencia_Suspencao" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovAdvertenciaSuspencao = New adodb.Recordset
    rsMovAdvertenciaSuspencao.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovAdvertenciaSuspencao.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovAdvertenciaSuspencao.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovAdvertenciaSuspencao("Empresa").Value
    Data = rsMovAdvertenciaSuspencao("Data").Value
    CodigoFuncionario = rsMovAdvertenciaSuspencao("Codigo do Funcionario").Value
    AdvertenciaouSuspencao = rsMovAdvertenciaSuspencao("Advertencia ou Suspencao").Value
    Dia = rsMovAdvertenciaSuspencao("Dia").Value
    Motivo1 = rsMovAdvertenciaSuspencao("Motivo_1").Value
    Motivo2 = rsMovAdvertenciaSuspencao("Motivo_2").Value
    Motivo3 = rsMovAdvertenciaSuspencao("Motivo_3").Value
    Motivo4 = rsMovAdvertenciaSuspencao("Motivo_4").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

