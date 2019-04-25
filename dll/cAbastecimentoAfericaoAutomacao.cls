VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAbastecimentoAfericaoAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Private mEmpresa As Integer
    Private mBico As Integer
    Private mDataAfericao As Date
    Private mPeriodo As Integer
    Private mDataAbastecimento As Date
    Private mHoraAbastecimento As Date
    Private mCodigoProduto As Integer
    Private mTipoCombustivel As String
    Private mValorUnitario As Currency
    Private mQuantidade As Currency
    Private mValorTotal As Currency
    Private mCodigoFuncionario As Integer

    Private rsAbastecimentoAfericao As New adodb.Recordset
    
    
    Public Property Get CodigoFuncionario() As Integer
        CodigoFuncionario = mCodigoFuncionario
    End Property
    Public Property Let CodigoFuncionario(ByVal Valor As Integer)
        mCodigoFuncionario = Valor
    End Property
    Public Property Get ValorTotal() As Currency
        ValorTotal = mValorTotal
    End Property
    Public Property Let ValorTotal(ByVal Valor As Currency)
        mValorTotal = Valor
    End Property
    Public Property Get Quantidade() As Currency
           Quantidade = mQuantidade
    End Property
    Public Property Let Quantidade(ByVal Valor As Currency)
           mQuantidade = Valor
    End Property
    Public Property Get ValorUnitario() As Currency
            ValorUnitario = mValorUnitario
    End Property
    Public Property Let ValorUnitario(ByVal Valor As Currency)
            mValorUnitario = Valor
    End Property
    Public Property Get TipoCombustivel() As String
            TipoCombustivel = mTipoCombustivel
    End Property
    Public Property Let TipoCombustivel(ByVal Valor As String)
            mTipoCombustivel = Valor
    End Property
    Public Property Get CodigoProduto() As Integer
            CodigoProduto = mCodigoProduto
    End Property
    Public Property Let CodigoProduto(ByVal Valor As Integer)
            mCodigoProduto = Valor
    End Property
    Public Property Get HoraAbastecimento() As Date
            HoraAbastecimento = mHoraAbastecimento
    End Property
    Public Property Let HoraAbastecimento(ByVal Valor As Date)
            mHoraAbastecimento = Valor
    End Property
    Public Property Get DataAbastecimento() As Date
           DataAbastecimento = mDataAbastecimento
    End Property
    Public Property Let DataAbastecimento(ByVal Valor As Date)
           mDataAbastecimento = Valor
    End Property
    Public Property Get Periodo() As Integer
           Periodo = mPeriodo
    End Property
    Public Property Let Periodo(ByVal Valor As Integer)
           mPeriodo = Valor
    End Property
    Public Property Get DataAfericao() As Date
           DataAfericao = mDataAfericao
    End Property
    Public Property Let DataAfericao(ByVal Valor As Date)
           mDataAfericao = Valor
    End Property
    Public Property Get Bico() As Integer
            Bico = mBico
    End Property
    Public Property Let Bico(ByVal Valor As Integer)
            mBico = Valor
    End Property
    Public Property Get Empresa() As Integer
           Empresa = mEmpresa
    End Property
    Public Property Let Empresa(ByVal Valor As Integer)
           mEmpresa = Valor
    End Property

Public Function Incluir() As Boolean
    Dim lRecordsAffected As Long

    On Error GoTo Err_Incluir
        Incluir = False
            
        gSQL = "INSERT INTO AbastecimentoAfericaoAutomacao ( Empresa, Bico, [Data da Afericao],Periodo,[Data do Abastecimento], [Hora do Abastecimento], "
        gSQL = gSQL & "[Codigo do Produto], [Tipo de Combustivel], [Valor Unitario], Quantidade, [Valor Total], [Codigo do Funcionario]) VALUES ( "
        Call sqlNumero(1, CStr(mEmpresa), ", ")
        Call sqlNumero(1, CStr(mBico), ", ")
        Call sqlData(1, CStr(mDataAfericao), ", ")
        Call sqlNumero(1, mPeriodo, ", ")
        Call sqlData(1, CStr(mDataAbastecimento), ", ")
        Call sqlHora(1, CStr(mHoraAbastecimento), ", ")
        Call sqlNumero(1, CStr(mCodigoProduto), ", ")
        Call sqlTexto(1, mTipoCombustivel, ", ")
        Call sqlValor(1, CStr(mValorUnitario), ", ")
        Call sqlValor(1, CStr(mQuantidade), ", ")
        Call sqlValor(1, CStr(mValorTotal), ", ")
        Call sqlNumero(1, mCodigoFuncionario, " )")
        
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
            
        If lRecordsAffected > 0 Then
            Incluir = True
        End If
        
        Exit Function
Err_Incluir:

End Function

Public Function Excluir(ByVal pEmpresa As Integer, ByVal pBico As Integer, ByVal pDataAfericao As Date, ByVal pDataAbastecimento As Date, ByVal pHora As Date, ByVal pCodigoFuncionario As Integer) As Boolean
    Dim lRecordsAffected As Long

    On Error GoTo Err_Excluir
    
        Excluir = False
        gSQL = "DELETE FROM AbastecimentoAfericaoAutomacao"
        gSQL = gSQL & " WHERE Empresa = " & pEmpresa
        gSQL = gSQL & " AND Bico = " & pBico
        gSQL = gSQL & " AND [Data da Afericao] = " & preparaData(pDataAfericao)
        gSQL = gSQL & " AND [Data do Abastecimento] = " & preparaData(pDataAbastecimento)
        gSQL = gSQL & " AND [Hora do Abastecimento] = " & preparaHora(pHora)
        gSQL = gSQL & " AND [Codigo do Funcionario] = " & pCodigoFuncionario
        
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        
        If lRecordsAffected > 0 Then
            Excluir = True
        End If
        Exit Function
        
Err_Excluir:

End Function

Private Sub AtribuiValor()

Dim xLocal As Integer
    xLocal = 1
    Empresa = rsAbastecimentoAfericao("Empresa").Value
    xLocal = 2
    Bico = rsAbastecimentoAfericao("Bico").Value
    xLocal = 3
    DataAfericao = rsAbastecimentoAfericao("Data da Afericao").Value
    xLocal = 4
    Periodo = rsAbastecimentoAfericao("Periodo").Value
    xLocal = 5
    DataAbastecimento = rsAbastecimentoAfericao("Data do Abastecimento").Value
    xLocal = 6
    HoraAbastecimento = rsAbastecimentoAfericao("Hora do Abastecimento").Value
    xLocal = 7
    CodigoProduto = rsAbastecimentoAfericao("Codigo do Produto").Value
    xLocal = 8
    TipoCombustivel = rsAbastecimentoAfericao("Tipo de Combustivel").Value
    xLocal = 9
    ValorUnitario = rsAbastecimentoAfericao("Valor Unitario").Value
    xLocal = 10
    Quantidade = rsAbastecimentoAfericao("Quantidade").Value
    xLocal = 11
    ValorTotal = rsAbastecimentoAfericao("Valor Total").Value
    xLocal = 12
    CodigoFuncionario = rsAbastecimentoAfericao("Codigo do Funcionario").Value
End Sub

Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsAbastecimentoAfericao = New adodb.Recordset
    rsAbastecimentoAfericao.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsAbastecimentoAfericao.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsAbastecimentoAfericao.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Function PreparaSQL(ByVal pCondicao As String, ByVal pTipoOrdem As String) As String
   PreparaSQL = Trim("SELECT Empresa, Bico, [Data da Afericao], Periodo, [Data da Abastecimento], [Hora do Abastecimento], [Codigo do Produto], [Tipo de Combustivel],[Valor Unitario], Quantidade, [Valor Total], [Codigo do Funcionario]  FROM AbastecimentoAfericaoAutomacao" & " " & pCondicao & " " & pTipoOrdem)
End Function

