VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLivroLMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mTipoCombustivel As String
Private mTipoLivro As String
Private mDataInicial As Date
Private mDataFinal As Date
Private mPaginaInicial As Integer
Private mPaginaFinal As Integer
Private mNumeroLivro As String
Private mConcluido As Boolean

Private rsLivroLMC As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get TipoCombustivel() As String
    TipoCombustivel = mTipoCombustivel
End Property
Public Property Let TipoCombustivel(ByVal Valor As String)
    mTipoCombustivel = Valor
End Property


Public Property Get TipoLivro() As String
    TipoLivro = mTipoLivro
End Property
Public Property Let TipoLivro(ByVal Valor As String)
    mTipoLivro = Valor
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


Public Property Get PaginaInicial() As Integer
    PaginaInicial = mPaginaInicial
End Property
Public Property Let PaginaInicial(ByVal Valor As Integer)
    mPaginaInicial = Valor
End Property


Public Property Get PaginaFinal() As Integer
    PaginaFinal = mPaginaFinal
End Property
Public Property Let PaginaFinal(ByVal Valor As Integer)
    mPaginaFinal = Valor
End Property


Public Property Get NumeroLivro() As String
    NumeroLivro = mNumeroLivro
End Property
Public Property Let NumeroLivro(ByVal Valor As String)
    mNumeroLivro = Valor
End Property


Public Property Get Concluido() As Boolean
    Concluido = mConcluido
End Property
Public Property Let Concluido(ByVal Valor As Boolean)
    mConcluido = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pTipoCombustivel As String, ByVal pTipoLivro As String, ByVal pDataInicial As Date) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    xCondicao = xCondicao & " AND [Tipo de Livro] = " & preparaTexto(pTipoLivro)
    xCondicao = xCondicao & " AND [Data Inicial] = " & preparaData(pDataInicial)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsLivroLMC.Close
    Set rsLivroLMC = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarCombustivelDataAproximada(ByVal pEmpresa As Integer, ByVal pTipoCombustivel As String, ByVal pTipoLivro As String, ByVal pData As Date) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCombustivelDataAproximada = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    xCondicao = xCondicao & " AND [Tipo de Livro] = " & preparaTexto(pTipoLivro)
    xCondicao = xCondicao & " AND [Data Inicial] <= " & preparaData(pData)
    xCondicao = xCondicao & " AND [Data Final] >= " & preparaData(pData)
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        LocalizarCombustivelDataAproximada = True
        AtribuiValor
    End If
    rsLivroLMC.Close
    Set rsLivroLMC = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarCombustivelConcluido(ByVal pEmpresa As Integer, ByVal pTipoCombustivel As String, ByVal pData As Date) As String
Dim xCondicao As String
Dim i As Integer

On Error GoTo trata_erro

    LocalizarCombustivelConcluido = "**"
    i = 1
    For i = 1 To 2
        gSQL = "SELECT Concluido"
        gSQL = gSQL & " FROM LivroLMC"
        gSQL = gSQL & " WHERE Empresa = " & pEmpresa
        If pTipoCombustivel <> "**" Then
            If i = 1 Then
                gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
            Else
                gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto("TC")
            End If
        End If
        gSQL = gSQL & " AND [Tipo de Livro] = " & preparaTexto("L")
        gSQL = gSQL & " AND [Data Inicial] <= " & preparaData(pData)
        gSQL = gSQL & " AND [Data Final] >= " & preparaData(pData)
        gSQL = gSQL & " AND MONTH([Data Inicial]) = " & Month(pData)
        gSQL = gSQL & " AND YEAR([Data Inicial]) = " & Year(pData)
        If Localizar(1) Then
            If rsLivroLMC("Concluido").Value = True Then
                LocalizarCombustivelConcluido = "SIM"
            Else
                LocalizarCombustivelConcluido = "NAO"
            End If
            Exit For
        Else
            'Verificação para lmc Anual
            rsLivroLMC.Close
            Set rsLivroLMC = Nothing
            gSQL = "SELECT Concluido"
            gSQL = gSQL & " FROM LivroLMC"
            gSQL = gSQL & " WHERE Empresa = " & pEmpresa
            If pTipoCombustivel <> "**" Then
                If i = 1 Then
                    gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
                Else
                    gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto("TC")
                End If
            End If
            gSQL = gSQL & " AND [Tipo de Livro] = " & preparaTexto("L")
            gSQL = gSQL & " AND [Data Inicial] <= " & preparaData(pData)
            gSQL = gSQL & " AND [Data Final] >= " & preparaData(pData)
            gSQL = gSQL & " AND MONTH([Data Inicial]) = " & 1
            gSQL = gSQL & " AND MONTH([Data Final]) = " & 12
            gSQL = gSQL & " AND YEAR([Data Inicial]) = " & Year(pData)
            If Localizar(1) Then
                If rsLivroLMC("Concluido").Value = True Then
                    LocalizarCombustivelConcluido = "SIM"
                Else
                    LocalizarCombustivelConcluido = "NAO"
                End If
                Exit For
            End If
        End If
    Next
    rsLivroLMC.Close
    Set rsLivroLMC = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO LivroLMC ( Empresa, [Tipo de Combustivel], [Tipo de Livro], [Data Inicial], [Data Final], "
    gSQL = gSQL & "[Pagina Inicial], [Pagina Final], [Numero do Livro], Concluido ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlTexto(1, mTipoCombustivel, ", ")
    Call sqlTexto(1, mTipoLivro, ", ")
    Call sqlData(1, mDataInicial, ", ")
    Call sqlData(1, mDataFinal, ", ")
    Call sqlNumero(1, mPaginaInicial, ", ")
    Call sqlNumero(1, mPaginaFinal, ", ")
    Call sqlTexto(1, mNumeroLivro, ", ")
    Call sqlBoolean(1, mConcluido, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pTipoCombustivel As String, ByVal pTipoLivro As String, ByVal pDataInicial As Date) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE LivroLMC SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlTexto(2, ", [Tipo de Combustivel] = ", mTipoCombustivel)
    Call sqlTexto(2, ", [Tipo de Livro] = ", mTipoLivro)
    Call sqlData(2, ", [Data Inicial] = ", mDataInicial)
    Call sqlData(2, ", [Data Final] = ", mDataFinal)
    Call sqlNumero(2, ", [Pagina Inicial] = ", mPaginaInicial)
    Call sqlNumero(2, ", [Pagina Final] = ", mPaginaFinal)
    Call sqlTexto(2, ", [Numero do Livro] = ", mNumeroLivro)
    Call sqlBoolean(2, ", [Concluido] = ", mConcluido)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    gSQL = gSQL & " AND [Tipo de Livro] = " & preparaTexto(pTipoLivro)
    gSQL = gSQL & " AND [Data Inicial] = " & preparaData(pDataInicial)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pTipoCombustivel As String, ByVal pTipoLivro As String, ByVal pDataInicial As Date) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM LivroLMC"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    gSQL = gSQL & " AND [Tipo de Livro] = " & preparaTexto(pTipoLivro)
    gSQL = gSQL & " AND [Data Inicial] = " & preparaData(pDataInicial)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Tipo de Combustivel], [Tipo de Livro], [Data Inicial], [Data Final], [Pagina Inicial], [Pagina Final], [Numero do Livro], Concluido FROM LivroLMC" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsLivroLMC = New adodb.Recordset
    rsLivroLMC.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsLivroLMC.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsLivroLMC.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsLivroLMC("Empresa").Value
    TipoCombustivel = rsLivroLMC("Tipo de Combustivel").Value
    TipoLivro = rsLivroLMC("Tipo de Livro").Value
    DataInicial = rsLivroLMC("Data Inicial").Value
    DataFinal = rsLivroLMC("Data Final").Value
    PaginaInicial = rsLivroLMC("Pagina Inicial").Value
    PaginaFinal = rsLivroLMC("Pagina Final").Value
    NumeroLivro = rsLivroLMC("Numero do Livro").Value
    Concluido = rsLivroLMC("Concluido").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

