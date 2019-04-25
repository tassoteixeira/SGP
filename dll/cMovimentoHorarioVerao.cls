VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoHorarioVerao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mCodigoECF As Integer
Private mDataParaInicioBloqueio As Date
Private mHoraParaInicioBloqueio As Date
Private mDataParaImpressaoReducaoZ As Date
Private mHoraParaImpressaoReducaoZ As Date
Private mComandoReducaoZConcluido As Boolean
Private mDataParaMudancaHorario As Date
Private mHoraParaMudancaHorario As Date
Private mEntradaHorarioVerao As Boolean
Private mComandoVeraoConcluido As Boolean

Private rsMovHorarioVerao As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get CodigoECF() As Integer
    CodigoECF = mCodigoECF
End Property
Public Property Let CodigoECF(ByVal Valor As Integer)
    mCodigoECF = Valor
End Property


Public Property Get DataParaInicioBloqueio() As Date
    DataParaInicioBloqueio = mDataParaInicioBloqueio
End Property
Public Property Let DataParaInicioBloqueio(ByVal Valor As Date)
    mDataParaInicioBloqueio = Valor
End Property


Public Property Get HoraParaInicioBloqueio() As Date
    HoraParaInicioBloqueio = mHoraParaInicioBloqueio
End Property
Public Property Let HoraParaInicioBloqueio(ByVal Valor As Date)
    mHoraParaInicioBloqueio = Valor
End Property


Public Property Get DataParaImpressaoReducaoZ() As Date
    DataParaImpressaoReducaoZ = mDataParaImpressaoReducaoZ
End Property
Public Property Let DataParaImpressaoReducaoZ(ByVal Valor As Date)
    mDataParaImpressaoReducaoZ = Valor
End Property


Public Property Get HoraParaImpressaoReducaoZ() As Date
    HoraParaImpressaoReducaoZ = mHoraParaImpressaoReducaoZ
End Property
Public Property Let HoraParaImpressaoReducaoZ(ByVal Valor As Date)
    mHoraParaImpressaoReducaoZ = Valor
End Property


Public Property Get ComandoReducaoZConcluido() As Boolean
    ComandoReducaoZConcluido = mComandoReducaoZConcluido
End Property
Public Property Let ComandoReducaoZConcluido(ByVal Valor As Boolean)
    mComandoReducaoZConcluido = Valor
End Property


Public Property Get DataParaMudancaHorario() As Date
    DataParaMudancaHorario = mDataParaMudancaHorario
End Property
Public Property Let DataParaMudancaHorario(ByVal Valor As Date)
    mDataParaMudancaHorario = Valor
End Property


Public Property Get HoraParaMudancaHorario() As Date
    HoraParaMudancaHorario = mHoraParaMudancaHorario
End Property
Public Property Let HoraParaMudancaHorario(ByVal Valor As Date)
    mHoraParaMudancaHorario = Valor
End Property


Public Property Get EntradaHorarioVerao() As Boolean
    EntradaHorarioVerao = mEntradaHorarioVerao
End Property
Public Property Let EntradaHorarioVerao(ByVal Valor As Boolean)
    mEntradaHorarioVerao = Valor
End Property


Public Property Get ComandoVeraoConcluido() As Boolean
    ComandoVeraoConcluido = mComandoVeraoConcluido
End Property
Public Property Let ComandoVeraoConcluido(ByVal Valor As Boolean)
    mComandoVeraoConcluido = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pDataParaInicioBloqueio As Date, ByVal pHoraParaInicioBloqueio As Date) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Codigo do ECF] = " & pCodigoECF
    xCondicao = xCondicao & " AND [Data para Inicio de Bloqueio] = " & preparaData(pDataParaInicioBloqueio)
    xCondicao = xCondicao & " AND [Hora para Inicio de Bloqueio] = " & preparaHora(pHoraParaInicioBloqueio)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsMovHorarioVerao.Close
    Set rsMovHorarioVerao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarCodigoPendente(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigoPendente = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Codigo do ECF] = " & pCodigoECF
    xCondicao = xCondicao & " AND [Comando Verao Concluido] = " & preparaBooleano(False)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigoPendente = Localizar(1)
    If LocalizarCodigoPendente = True Then
        AtribuiValor
    End If
    rsMovHorarioVerao.Close
    Set rsMovHorarioVerao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Codigo do ECF] < " & mCodigoECF
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do ECF] DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsMovHorarioVerao.Close
    Set rsMovHorarioVerao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do ECF] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsMovHorarioVerao.Close
    Set rsMovHorarioVerao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Codigo do ECF] > " & mCodigoECF
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do ECF] ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsMovHorarioVerao.Close
    Set rsMovHorarioVerao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do ECF] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsMovHorarioVerao.Close
    Set rsMovHorarioVerao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO MovimentoHorarioVerao ( Empresa, [Codigo do ECF], [Data para Inicio de Bloqueio], [Hora para Inicio de Bloqueio], [Data para Impressao da Reducao Z], "
    gSQL = gSQL & "[Hora para Impressao da Reducao Z], [Comando Reducao ZConcluido], [Data para Mudanca de Horario], [Hora para Mudanca de Horario], [Entrada de Horario de Verao], [Comando Verao Concluido] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mCodigoECF, ", ")
    Call sqlData(1, mDataParaInicioBloqueio, ", ")
    Call sqlHora(1, mHoraParaInicioBloqueio, ", ")
    Call sqlData(1, mDataParaImpressaoReducaoZ, ", ")
    Call sqlHora(1, mHoraParaImpressaoReducaoZ, ", ")
    Call sqlBoolean(1, mComandoReducaoZConcluido, ", ")
    Call sqlData(1, mDataParaMudancaHorario, ", ")
    Call sqlHora(1, mHoraParaMudancaHorario, ", ")
    Call sqlBoolean(1, mEntradaHorarioVerao, ", ")
    Call sqlBoolean(1, mComandoVeraoConcluido, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pDataParaInicioBloqueio As Date, ByVal pHoraParaInicioBloqueio As Date) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE MovimentoHorarioVerao SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", [Codigo do ECF] = ", mCodigoECF)
    Call sqlData(2, ", [Data para Inicio de Bloqueio] = ", mDataParaInicioBloqueio)
    Call sqlHora(2, ", [Hora para Inicio de Bloqueio] = ", mHoraParaInicioBloqueio)
    Call sqlData(2, ", [Data para Impressao da Reducao Z] = ", mDataParaImpressaoReducaoZ)
    Call sqlHora(2, ", [Hora para Impressao da Reducao Z] = ", mHoraParaImpressaoReducaoZ)
    Call sqlBoolean(2, ", [Comando Reducao Z Concluido] = ", mComandoReducaoZConcluido)
    Call sqlData(2, ", [Data para Mudanca de Horario] = ", mDataParaMudancaHorario)
    Call sqlHora(2, ", [Hora para Mudanca de Horario] = ", mHoraParaMudancaHorario)
    Call sqlBoolean(2, ", [Entrada de Horario de Verao] = ", mEntradaHorarioVerao)
    Call sqlBoolean(2, ", [Comando Verao Concluido] = ", mComandoVeraoConcluido)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo do ECF] = " & pCodigoECF
    gSQL = gSQL & " AND [Data para Inicio de Bloqueio] = " & preparaData(pDataParaInicioBloqueio)
    gSQL = gSQL & " AND [Hora para Inicio de Bloqueio] = " & preparaHora(pHoraParaInicioBloqueio)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pCodigoECF As Integer, ByVal pDataParaInicioBloqueio As Date, ByVal pHoraParaInicioBloqueio As Date) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM MovimentoHorarioVerao"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo do ECF] = " & pCodigoECF
    gSQL = gSQL & " AND [Data para Inicio de Bloqueio] = " & preparaData(pDataParaInicioBloqueio)
    gSQL = gSQL & " AND [Hora para Inicio de Bloqueio] = " & preparaHora(pHoraParaInicioBloqueio)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Codigo do ECF], [Data para Inicio de Bloqueio], [Hora para Inicio de Bloqueio], [Data para Impressao da Reducao Z], [Hora para Impressao da Reducao Z], [Comando Reducao Z Concluido], [Data para Mudanca de Horario], [Hora para Mudanca de Horario], [Entrada de Horario de Verao], [Comando Verao Concluido] FROM MovimentoHorarioVerao" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovHorarioVerao = New adodb.Recordset
    rsMovHorarioVerao.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovHorarioVerao.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovHorarioVerao.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovHorarioVerao("Empresa").Value
    CodigoECF = rsMovHorarioVerao("Codigo do ECF").Value
    DataParaInicioBloqueio = rsMovHorarioVerao("Data para Inicio de Bloqueio").Value
    HoraParaInicioBloqueio = rsMovHorarioVerao("Hora para Inicio de Bloqueio").Value
    If IsNull(rsMovHorarioVerao("Data para Impressao da Reducao Z").Value) Then
        DataParaImpressaoReducaoZ = "00:00:00"
    Else
        DataParaImpressaoReducaoZ = rsMovHorarioVerao("Data para Impressao da Reducao Z").Value
    End If
    If IsNull(rsMovHorarioVerao("Hora para Impressao da Reducao Z").Value) Then
        HoraParaImpressaoReducaoZ = "00:00:00"
    Else
        HoraParaImpressaoReducaoZ = rsMovHorarioVerao("Hora para Impressao da Reducao Z").Value
    End If
    ComandoReducaoZConcluido = rsMovHorarioVerao("Comando Reducao Z Concluido").Value
    DataParaMudancaHorario = rsMovHorarioVerao("Data para Mudanca de Horario").Value
    HoraParaMudancaHorario = rsMovHorarioVerao("Hora para Mudanca de Horario").Value
    EntradaHorarioVerao = rsMovHorarioVerao("Entrada de Horario de Verao").Value
    ComandoVeraoConcluido = rsMovHorarioVerao("Comando Verao Concluido").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

