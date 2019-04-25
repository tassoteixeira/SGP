VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPremiacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mMesAno As Date
Private mValorBase As Currency
Private mPercentualBase1 As Currency
Private mPercentualComissao1 As Currency
Private mPercentualBase2 As Currency
Private mPercentualComissao2 As Currency
Private mPercentualBase3 As Currency
Private mPercentualComissao3 As Currency

Private rsPremiacao As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get MesAno() As Date
    MesAno = mMesAno
End Property
Public Property Let MesAno(ByVal Valor As Date)
    mMesAno = Valor
End Property


Public Property Get ValorBase() As Currency
    ValorBase = mValorBase
End Property
Public Property Let ValorBase(ByVal Valor As Currency)
    mValorBase = Valor
End Property


Public Property Get PercentualBase1() As Currency
    PercentualBase1 = mPercentualBase1
End Property
Public Property Let PercentualBase1(ByVal Valor As Currency)
    mPercentualBase1 = Valor
End Property


Public Property Get PercentualComissao1() As Currency
    PercentualComissao1 = mPercentualComissao1
End Property
Public Property Let PercentualComissao1(ByVal Valor As Currency)
    mPercentualComissao1 = Valor
End Property


Public Property Get PercentualBase2() As Currency
    PercentualBase2 = mPercentualBase2
End Property
Public Property Let PercentualBase2(ByVal Valor As Currency)
    mPercentualBase2 = Valor
End Property


Public Property Get PercentualComissao2() As Currency
    PercentualComissao2 = mPercentualComissao2
End Property
Public Property Let PercentualComissao2(ByVal Valor As Currency)
    mPercentualComissao2 = Valor
End Property


Public Property Get PercentualBase3() As Currency
    PercentualBase3 = mPercentualBase3
End Property
Public Property Let PercentualBase3(ByVal Valor As Currency)
    mPercentualBase3 = Valor
End Property


Public Property Get PercentualComissao3() As Currency
    PercentualComissao3 = mPercentualComissao3
End Property
Public Property Let PercentualComissao3(ByVal Valor As Currency)
    mPercentualComissao3 = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nMesAno As Date) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Mes Ano] = " & preparaData(nMesAno)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Mes Ano] < " & preparaData(mMesAno)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, [Mes Ano] DESC")
    LocalizarAnterior = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Mes Ano] > " & preparaData(CDate("01/01/1900"))
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Mes Ano] ASC")
    LocalizarPrimeiro = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Mes Ano] > " & preparaData(mMesAno)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, [Mes Ano] ASC")
    LocalizarProximo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Mes Ano] < " & preparaData("31/12/2500")
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, [Mes Ano] DESC")
    LocalizarUltimo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Tabela_Premiacao ( Empresa, [Mes Ano], [Valor Base], [Percentual Base 1], [Percentual Comissao 1], "
    gSQL = gSQL & "[Percentual Base 2], [Percentual Comissao 2], [Percentual Base 3], [Percentual Comissao 3] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mMesAno, ", ")
    Call sqlValor(1, mValorBase, ", ")
    Call sqlValor(1, mPercentualBase1, ", ")
    Call sqlValor(1, mPercentualComissao1, ", ")
    Call sqlValor(1, mPercentualBase2, ", ")
    Call sqlValor(1, mPercentualComissao2, ", ")
    Call sqlValor(1, mPercentualBase3, ", ")
    Call sqlValor(1, mPercentualComissao3, " )")
    gConn.Execute gSQL, Incluir, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nMesAno As Date) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Tabela_Premiacao SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", [Mes Ano] = ", mMesAno)
    Call sqlValor(2, ", [Valor Base] = ", mValorBase)
    Call sqlValor(2, ", [Percentual Base 1] = ", mPercentualBase1)
    Call sqlValor(2, ", [Percentual Comissao 1] = ", mPercentualComissao1)
    Call sqlValor(2, ", [Percentual Base 2] = ", mPercentualBase2)
    Call sqlValor(2, ", [Percentual Comissao 2] = ", mPercentualComissao2)
    Call sqlValor(2, ", [Percentual Base 3] = ", mPercentualBase3)
    Call sqlValor(2, ", [Percentual Comissao 3] = ", mPercentualComissao3)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Mes Ano] = " & preparaData(nMesAno)
    gConn.Execute gSQL, Alterar, adCmdText + adExecuteNoRecords
    Alterar = True
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nMesAno As Date) As Boolean

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Tabela_Premiacao"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Mes Ano] = " & preparaData(nMesAno)
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Mes Ano], [Valor Base], [Percentual Base 1], [Percentual Comissao 1], [Percentual Base 2], [Percentual Comissao 2], [Percentual Base 3], [Percentual Comissao 3] FROM Tabela_Premiacao" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsPremiacao = New adodb.Recordset
    rsPremiacao.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsPremiacao.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsPremiacao.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsPremiacao("Empresa").Value
    MesAno = rsPremiacao("Mes Ano").Value
    ValorBase = rsPremiacao("Valor Base").Value
    PercentualBase1 = rsPremiacao("Percentual Base 1").Value
    PercentualComissao1 = rsPremiacao("Percentual Comissao 1").Value
    PercentualBase2 = rsPremiacao("Percentual Base 2").Value
    PercentualComissao2 = rsPremiacao("Percentual Comissao 2").Value
    PercentualBase3 = rsPremiacao("Percentual Base 3").Value
    PercentualComissao3 = rsPremiacao("Percentual Comissao 3").Value
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsPremiacao.Close
    Set rsPremiacao = Nothing
End Function

