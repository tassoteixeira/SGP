VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFinSaldoTesouraria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    Private mEmpresa As Integer
    Private mData As Date
    Private mCodigoTipoMovimento As Integer
    Private mSaldo As Currency
    
    Private rsSaldoTesouraria As New adodb.Recordset

Public Function AlterarSaldo(ByVal pEmpresa As Integer, ByVal pCodigoTipoMovimento As Integer, ByVal pData As Date, ByVal pValor As Currency, ByVal pAumentar As Boolean) As Boolean
    Dim lRecordsAffected As Long

    On Error GoTo Err_Alterar
        AlterarSaldo = False

        Call ExisteCriaSaldoData(pEmpresa, pCodigoTipoMovimento, pData)
        
            gSQL = "UPDATE FinSaldoTesouraria SET "
            If pAumentar Then
                Call sqlValor(2, "Saldo =  Saldo + ", pValor)
            Else
                Call sqlValor(2, "Saldo = Saldo - ", pValor)
            End If
            gSQL = gSQL & " WHERE Empresa = " & pEmpresa
            gSQL = gSQL & " AND [Codigo do Tipo de Movimento] = " & pCodigoTipoMovimento
            gSQL = gSQL & " AND Data >= " & preparaData(pData)
            
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
            
        If lRecordsAffected > 0 Then
            AlterarSaldo = True
        Else
            Call CriaLogCadastroDll2("cFinSaldoTesouraria:AlterarSaldo - Erro ao alterar registro.", Err.Description, gSQL)
        End If
        Exit Function

Err_Alterar:
    Call CriaLogCadastroDll2("cFinSaldoTesouraria:AlterarSaldo - Erro não identificado.", Err.Description, gSQL)
    Exit Function
End Function
Public Function Incluir(ByVal pEmpresa As Integer) As Boolean
    Dim lRecordsAffected As Long

    On Error GoTo Err_Incluir
        Incluir = False
           
           gSQL = "INSERT INTO FinSaldoTesouraria ( Empresa, Data, [Codigo do Tipo de Movimento], Saldo  ) VALUES ( "
            Call sqlNumero(1, pEmpresa, ", ")
            Call sqlData(1, mData, ", ")
            Call sqlNumero(1, mCodigoTipoMovimento, ", ")
            Call sqlValor(1, mSaldo, " )")
        
            gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
            If lRecordsAffected > 0 Then
                Incluir = True
            End If
        Exit Function
Err_Incluir:
    Call CriaLogCadastroDll2("cFinSaldoTesouraria:Incluir - Erro não identificado.", Err.Description, gSQL)
    Exit Function
End Function

Public Function ExisteCriaSaldoData(ByVal pEmpresa As Integer, ByVal pCodigoTipoMovimento As Integer, ByVal pData As Date) As Boolean
        Dim xCriaSaldo As Boolean
        Dim xSaldoInicial As Currency
        
        xCriaSaldo = False
        xSaldoInicial = 0
        
        ExisteCriaSaldoData = False
        
        gSQL = "SELECT Data, Saldo"
        gSQL = gSQL & " FROM FinSaldoTesouraria"
        gSQL = gSQL & " WHERE Empresa = " & pEmpresa
        gSQL = gSQL & "   AND [Codigo do Tipo de Movimento] = " & pCodigoTipoMovimento
        gSQL = gSQL & "   AND Data <= " & preparaData(pData)
        gSQL = gSQL & " ORDER BY Data DESC"
        
        
        
        If Localizar(1) Then
            If rsSaldoTesouraria("Data").Value < pData Then
                    xSaldoInicial = rsSaldoTesouraria("Saldo").Value
                    xCriaSaldo = True
            Else
                xCriaSaldo = True
            End If
        Else
            xCriaSaldo = True
        End If
        
        Call CriaLogCadastroDll2("cFinSaldoTesouraria:ExisteCriaSaldoData - xCriaSaldo 2." & xCriaSaldo, "", "")

        If xCriaSaldo Then
            mCodigoTipoMovimento = pCodigoTipoMovimento
            mData = pData
            mSaldo = xSaldoInicial
            If Incluir(pEmpresa) Then
                ExisteCriaSaldoData = True
            End If
        Else
            ExisteCriaSaldoData = True
        End If
End Function

Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function
Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsSaldoTesouraria = New adodb.Recordset
    rsSaldoTesouraria.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsSaldoTesouraria.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsSaldoTesouraria.EOF Then
        AtualizaRecordset = True
    End If
End Function


