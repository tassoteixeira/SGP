VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrevisaoVendaPrazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mData As Date
Private mTipoCombustivel As String
Private mPrevisaoVendaPrazo As Currency
Private mMediaVendaDiariaPrazo As Currency
Private mTotalVenda As Currency
Private mQuantidadeVendaPrazo As Currency
Private mSaldo As Currency
Private mHora As Date

Private rsPrevisaoVendaPrazo As New adodb.Recordset


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


Public Property Get TipoCombustivel() As String
    TipoCombustivel = mTipoCombustivel
End Property
Public Property Let TipoCombustivel(ByVal Valor As String)
    mTipoCombustivel = Valor
End Property


Public Property Get PrevisaoVendaPrazo() As Currency
    PrevisaoVendaPrazo = mPrevisaoVendaPrazo
End Property
Public Property Let PrevisaoVendaPrazo(ByVal Valor As Currency)
    mPrevisaoVendaPrazo = Valor
End Property


Public Property Get MediaVendaDiariaPrazo() As Currency
    MediaVendaDiariaPrazo = mMediaVendaDiariaPrazo
End Property
Public Property Let MediaVendaDiariaPrazo(ByVal Valor As Currency)
    mMediaVendaDiariaPrazo = Valor
End Property


Public Property Get TotalVenda() As Currency
    TotalVenda = mTotalVenda
End Property
Public Property Let TotalVenda(ByVal Valor As Currency)
    mTotalVenda = Valor
End Property


Public Property Get QuantidadeVendaPrazo() As Currency
    QuantidadeVendaPrazo = mQuantidadeVendaPrazo
End Property
Public Property Let QuantidadeVendaPrazo(ByVal Valor As Currency)
    mQuantidadeVendaPrazo = Valor
End Property


Public Property Get Saldo() As Currency
    Saldo = mSaldo
End Property
Public Property Let Saldo(ByVal Valor As Currency)
    mSaldo = Valor
End Property


Public Property Get Hora() As Date
    Hora = mHora
End Property
Public Property Let Hora(ByVal Valor As Date)
    mHora = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nTipoCombustivel As String, ByVal nData As Date) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(nTipoCombustivel)
    xCondicao = xCondicao & " AND Data = " & preparaData(nData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Hora")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsPrevisaoVendaPrazo.Close
    Set rsPrevisaoVendaPrazo = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xRegIgual As Boolean
Dim xSQL As String
Dim xOrdem As String

On Error GoTo trata_erro

    LocalizarAnterior = False
    xRegIgual = False
    xSQL = " WHERE Empresa = " & mEmpresa
    'xSQL = xSQL & " AND [Tipo de Combustivel] = " & preparaTexto(mTipoCombustivel)
    xSQL = xSQL & " AND Data <= " & preparaData(mData)
    xOrdem = "ORDER BY Data DESC, [Tipo de Combustivel] DESC, Hora DESC"
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(30) Then
        If rsPrevisaoVendaPrazo.RecordCount > 1 Then
            Do Until rsPrevisaoVendaPrazo.EOF
                If xRegIgual = False Then
                    If rsPrevisaoVendaPrazo("Data").Value = mData And rsPrevisaoVendaPrazo("Tipo de Combustivel").Value = mTipoCombustivel And rsPrevisaoVendaPrazo("Hora").Value = mHora Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsPrevisaoVendaPrazo.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsPrevisaoVendaPrazo.Close
    Set rsPrevisaoVendaPrazo = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    'xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(mTipoCombustivel)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, [Tipo de Combustivel] DESC, Hora ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsPrevisaoVendaPrazo.Close
    Set rsPrevisaoVendaPrazo = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xRegIgual As Boolean
Dim xSQL As String
Dim xOrdem As String

On Error GoTo trata_erro

    LocalizarProximo = False
    xRegIgual = False
    xSQL = " WHERE Empresa = " & mEmpresa
    'xSQL = xSQL & " AND [Tipo de Combustivel] = " & preparaTexto(mTipoCombustivel)
    xSQL = xSQL & " AND Data >= " & preparaData(mData)
    xOrdem = "ORDER BY Data ASC, [Tipo de Combustivel] ASC, Hora ASC"
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(30) Then
        If rsPrevisaoVendaPrazo.RecordCount > 1 Then
            Do Until rsPrevisaoVendaPrazo.EOF
                If xRegIgual = False Then
                    If rsPrevisaoVendaPrazo("Data").Value = mData And rsPrevisaoVendaPrazo("Tipo de Combustivel").Value = mTipoCombustivel And rsPrevisaoVendaPrazo("Hora").Value = mHora Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsPrevisaoVendaPrazo.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsPrevisaoVendaPrazo.Close
    Set rsPrevisaoVendaPrazo = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, [Tipo de Combustivel] DESC, Hora DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsPrevisaoVendaPrazo.Close
    Set rsPrevisaoVendaPrazo = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Previsao_Venda_Prazo ( Empresa, Data, [Tipo de Combustivel], [Previsao de Venda a Prazo], [Media de Venda Diaria a Prazo], "
    gSQL = gSQL & "[Total da Venda], [Quantidade de Venda a Prazo], Saldo, Hora ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlTexto(1, mTipoCombustivel, ", ")
    Call sqlValor(1, mPrevisaoVendaPrazo, ", ")
    Call sqlValor(1, mMediaVendaDiariaPrazo, ", ")
    Call sqlValor(1, mTotalVenda, ", ")
    Call sqlValor(1, mQuantidadeVendaPrazo, ", ")
    Call sqlValor(1, mSaldo, ", ")
    Call sqlHora(1, mHora, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function ExcluiRegistros(ByVal nEmpresa As Integer, ByVal nTipoCombustivel As String, ByVal nData As Date) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    ExcluiRegistros = False
    gSQL = "DELETE FROM Previsao_Venda_Prazo"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto(nTipoCombustivel)
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        ExcluiRegistros = True
    End If
    Exit Function

Err_Excluir:
End Function


Public Function GravaVendaPrazoDia(ByVal nEmpresa As Integer, ByVal nTipoCombustivel As String, ByVal nData As Date, ByVal nQtdParaVendaPrazo As Currency, ByVal nQtdVendaCombustivel As Currency) As Currency
Dim xSQL As String
Dim xCondicao As String
Dim xSomaVendaDiariaPrazoDia As Currency
Dim xQtdRestanteDia As Currency
Dim xMediaVendaDiariaPrazo As Currency
Dim xPercentualVendaDiariaPrazo As Currency
Dim xSaldo As Currency

On Error GoTo Err_TotalVendaPrazoDia
    
    GravaVendaPrazoDia = 0
    xSomaVendaDiariaPrazoDia = 0
    xMediaVendaDiariaPrazo = 0
    xSaldo = 0
    
    'Busca [Media de Venda Diaria a Prazo]
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(nTipoCombustivel)
    xCondicao = xCondicao & " AND Data <= " & preparaData(nData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, Hora DESC")
    If Localizar(1) Then
        xMediaVendaDiariaPrazo = rsPrevisaoVendaPrazo("Media de Venda Diaria a Prazo").Value
        xPercentualVendaDiariaPrazo = rsPrevisaoVendaPrazo("Media de Venda Diaria a Prazo").Value
        xSaldo = rsPrevisaoVendaPrazo("Saldo").Value
    End If
    rsPrevisaoVendaPrazo.Close
    Set rsPrevisaoVendaPrazo = Nothing
    If xMediaVendaDiariaPrazo = 0 Or xSaldo = 0 Then
        Exit Function
    End If
    
    'Busca Soma do dia [Quantidade de Venda a Prazo]
    xSQL = ""
    xSQL = xSQL & "SELECT SUM([Quantidade de Venda a Prazo]) AS Total"
    xSQL = xSQL & "  FROM Previsao_Venda_Prazo"
    xSQL = xSQL & " WHERE Empresa = " & nEmpresa
    xSQL = xSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(nTipoCombustivel)
    xSQL = xSQL & "   AND Data = " & preparaData(nData)
    
    Set rsPrevisaoVendaPrazo = New adodb.Recordset
    rsPrevisaoVendaPrazo.CursorLocation = adUseClient
    rsPrevisaoVendaPrazo.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsPrevisaoVendaPrazo.EOF Then
        If IsNull(rsPrevisaoVendaPrazo("Total").Value) Then
            xSomaVendaDiariaPrazoDia = 0
        Else
            xSomaVendaDiariaPrazoDia = rsPrevisaoVendaPrazo("Total").Value
        End If
    End If
    rsPrevisaoVendaPrazo.Close
    Set rsPrevisaoVendaPrazo = Nothing
    
    'Atribui Valores para Gravação
    mEmpresa = nEmpresa
    mData = nData
    mTipoCombustivel = nTipoCombustivel
    mPrevisaoVendaPrazo = 0
    mMediaVendaDiariaPrazo = xMediaVendaDiariaPrazo
    mTotalVenda = nQtdVendaCombustivel
    mQuantidadeVendaPrazo = 0
    mSaldo = xSaldo
    mHora = Format(Time, "hh:mm:ss")
    
    'xQtdRestanteDia = xMediaVendaDiariaPrazo - xSomaVendaDiariaPrazoDia
    'If xQtdRestanteDia > 0 Then
    '    If (xQtdRestanteDia - nQtdParaVendaPrazo) >= 0 Then
    '        mQuantidadeVendaPrazo = nQtdParaVendaPrazo
    '    Else
    '        mQuantidadeVendaPrazo = xQtdRestanteDia
    '    End If
    '    If mSaldo < mQuantidadeVendaPrazo Then
    '        mQuantidadeVendaPrazo = mSaldo
    '    End If
    '    GravaVendaPrazoDia = mQuantidadeVendaPrazo
    '    mSaldo = mSaldo - mQuantidadeVendaPrazo
    '    If Not Incluir Then
    '        'MsgBox "erro"
    '    End If
    'End If
    xQtdRestanteDia = (nQtdVendaCombustivel * xMediaVendaDiariaPrazo / 100) - xSomaVendaDiariaPrazoDia
    



    If xQtdRestanteDia > 0 Then
        If (xQtdRestanteDia - nQtdParaVendaPrazo) >= 0 Then
            mQuantidadeVendaPrazo = nQtdParaVendaPrazo
        Else
            mQuantidadeVendaPrazo = xQtdRestanteDia
        End If
        If mSaldo < mQuantidadeVendaPrazo Then
            mQuantidadeVendaPrazo = mSaldo
        End If
        GravaVendaPrazoDia = mQuantidadeVendaPrazo
        mSaldo = mSaldo - mQuantidadeVendaPrazo
        If Not Incluir Then
            'MsgBox "erro"
        End If
    End If
    Exit Function
    
Err_TotalVendaPrazoDia:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, [Tipo de Combustivel], [Previsao de Venda a Prazo], [Media de Venda Diaria a Prazo], [Total da Venda], [Quantidade de Venda a Prazo], Saldo, Hora FROM Previsao_Venda_Prazo" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsPrevisaoVendaPrazo = New adodb.Recordset
    rsPrevisaoVendaPrazo.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsPrevisaoVendaPrazo.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsPrevisaoVendaPrazo.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsPrevisaoVendaPrazo("Empresa").Value
    Data = rsPrevisaoVendaPrazo("Data").Value
    TipoCombustivel = rsPrevisaoVendaPrazo("Tipo de Combustivel").Value
    PrevisaoVendaPrazo = rsPrevisaoVendaPrazo("Previsao de Venda a Prazo").Value
    MediaVendaDiariaPrazo = rsPrevisaoVendaPrazo("Media de Venda Diaria a Prazo").Value
    TotalVenda = rsPrevisaoVendaPrazo("Total da Venda").Value
    QuantidadeVendaPrazo = rsPrevisaoVendaPrazo("Quantidade de Venda a Prazo").Value
    Saldo = rsPrevisaoVendaPrazo("Saldo").Value
    Hora = rsPrevisaoVendaPrazo("Hora").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function


Public Function MediaAnterior(ByVal nEmpresa As Integer, ByVal nTipoCombustivel As String, ByVal nData As Date, ByVal nHora As Date) As Currency
Dim xSQL As String
Dim xOrdem As String

On Error GoTo Err_MediaAnterior

    MediaAnterior = 0
    
    'Busca Média Anterior
    xSQL = " WHERE Empresa = " & nEmpresa
    xSQL = xSQL & " AND [Tipo de Combustivel] = " & preparaTexto(nTipoCombustivel)
    xSQL = xSQL & " AND Data <= " & preparaData(nData)
    xOrdem = "ORDER BY Data DESC, Hora DESC"
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(20) Then
        If rsPrevisaoVendaPrazo.RecordCount > 1 Then
            Do Until rsPrevisaoVendaPrazo.EOF
                If rsPrevisaoVendaPrazo("Data").Value <> nData Or rsPrevisaoVendaPrazo("Hora").Value < nHora Then
                    MediaAnterior = rsPrevisaoVendaPrazo("Media de Venda Diaria a Prazo").Value
                    Exit Do
                End If
                rsPrevisaoVendaPrazo.MoveNext
            Loop
        End If
    End If
    rsPrevisaoVendaPrazo.Close
    Set rsPrevisaoVendaPrazo = Nothing
    Exit Function
    
Err_MediaAnterior:
End Function


Public Function SaldoAnterior(ByVal nEmpresa As Integer, ByVal nTipoCombustivel As String, ByVal nData As Date, ByVal nHora As Date) As Currency
Dim xSQL As String
Dim xOrdem As String

On Error GoTo Err_SaldoAnterior

    SaldoAnterior = 0
    
    'Busca Saldo Anterior
    xSQL = " WHERE Empresa = " & nEmpresa
    xSQL = xSQL & " AND [Tipo de Combustivel] = " & preparaTexto(nTipoCombustivel)
    xSQL = xSQL & " AND Data <= " & preparaData(nData)
    xOrdem = "ORDER BY Data DESC, Hora DESC"
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(20) Then
        If rsPrevisaoVendaPrazo.RecordCount > 1 Then
            Do Until rsPrevisaoVendaPrazo.EOF
                If rsPrevisaoVendaPrazo("Data").Value <> nData Or rsPrevisaoVendaPrazo("Hora").Value < nHora Then
                    SaldoAnterior = rsPrevisaoVendaPrazo("Saldo").Value
                    Exit Do
                End If
                rsPrevisaoVendaPrazo.MoveNext
            Loop
        End If
    End If
    rsPrevisaoVendaPrazo.Close
    Set rsPrevisaoVendaPrazo = Nothing
    Exit Function
    
Err_SaldoAnterior:
End Function


Public Function TotalVendaPrazoDia(ByVal nEmpresa As Integer, ByVal nTipoCombustivel As String, ByVal nData As Date) As Currency
Dim xSQL As String

On Error GoTo Err_TotalVendaPrazoDia

    TotalVendaPrazoDia = 0
    
    'Busca Soma do dia [Quantidade de Venda a Prazo]
    xSQL = ""
    xSQL = xSQL & "SELECT SUM([Quantidade de Venda a Prazo]) AS Total"
    xSQL = xSQL & "  FROM Previsao_Venda_Prazo"
    xSQL = xSQL & " WHERE Empresa = " & nEmpresa
    xSQL = xSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(nTipoCombustivel)
    xSQL = xSQL & "   AND Data = " & preparaData(nData)
    
    Set rsPrevisaoVendaPrazo = New adodb.Recordset
    rsPrevisaoVendaPrazo.CursorLocation = adUseClient
    rsPrevisaoVendaPrazo.Open xSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsPrevisaoVendaPrazo.EOF Then
        If IsNull(rsPrevisaoVendaPrazo("Total").Value) Then
            TotalVendaPrazoDia = 0
        Else
            TotalVendaPrazoDia = rsPrevisaoVendaPrazo("Total").Value
        End If
    End If
    rsPrevisaoVendaPrazo.Close
    Set rsPrevisaoVendaPrazo = Nothing
    Exit Function
    
Err_TotalVendaPrazoDia:
End Function


