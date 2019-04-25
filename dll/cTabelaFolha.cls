VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTabelaFolha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mMesAno As String
Private mSalarioMinimo As Currency
Private mCestaBasica As Currency
Private mSalarioFamilia1 As Currency
Private mSalarioFamiliaAcimade As Currency
Private mSalarioFamilia2 As Currency
Private mValorInicial1 As Currency
Private mValorFinal1 As Currency
Private mPercentual1 As Currency
Private mValorInicial2 As Currency
Private mValorFinal2 As Currency
Private mPercentual2 As Currency
Private mValorInicial3 As Currency
Private mValorFinal3 As Currency
Private mPercentual3 As Currency
Private mValorInicial4 As Currency
Private mValorFinal4 As Currency
Private mPercentual4 As Currency
Private mObservacao1 As String
Private mObservacao2 As String
Private mObservacao3 As String

Private rsTabelaFolha As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get MesAno() As String
    MesAno = mMesAno
End Property
Public Property Let MesAno(ByVal Valor As String)
    mMesAno = Valor
End Property


Public Property Get SalarioMinimo() As Currency
    SalarioMinimo = mSalarioMinimo
End Property
Public Property Let SalarioMinimo(ByVal Valor As Currency)
    mSalarioMinimo = Valor
End Property


Public Property Get CestaBasica() As Currency
    CestaBasica = mCestaBasica
End Property
Public Property Let CestaBasica(ByVal Valor As Currency)
    mCestaBasica = Valor
End Property


Public Property Get SalarioFamilia1() As Currency
    SalarioFamilia1 = mSalarioFamilia1
End Property
Public Property Let SalarioFamilia1(ByVal Valor As Currency)
    mSalarioFamilia1 = Valor
End Property


Public Property Get SalarioFamiliaAcimade() As Currency
    SalarioFamiliaAcimade = mSalarioFamiliaAcimade
End Property
Public Property Let SalarioFamiliaAcimade(ByVal Valor As Currency)
    mSalarioFamiliaAcimade = Valor
End Property


Public Property Get SalarioFamilia2() As Currency
    SalarioFamilia2 = mSalarioFamilia2
End Property
Public Property Let SalarioFamilia2(ByVal Valor As Currency)
    mSalarioFamilia2 = Valor
End Property


Public Property Get ValorInicial1() As Currency
    ValorInicial1 = mValorInicial1
End Property
Public Property Let ValorInicial1(ByVal Valor As Currency)
    mValorInicial1 = Valor
End Property


Public Property Get ValorFinal1() As Currency
    ValorFinal1 = mValorFinal1
End Property
Public Property Let ValorFinal1(ByVal Valor As Currency)
    mValorFinal1 = Valor
End Property


Public Property Get Percentual1() As Currency
    Percentual1 = mPercentual1
End Property
Public Property Let Percentual1(ByVal Valor As Currency)
    mPercentual1 = Valor
End Property


Public Property Get ValorInicial2() As Currency
    ValorInicial2 = mValorInicial2
End Property
Public Property Let ValorInicial2(ByVal Valor As Currency)
    mValorInicial2 = Valor
End Property


Public Property Get ValorFinal2() As Currency
    ValorFinal2 = mValorFinal2
End Property
Public Property Let ValorFinal2(ByVal Valor As Currency)
    mValorFinal2 = Valor
End Property


Public Property Get Percentual2() As Currency
    Percentual2 = mPercentual2
End Property
Public Property Let Percentual2(ByVal Valor As Currency)
    mPercentual2 = Valor
End Property


Public Property Get ValorInicial3() As Currency
    ValorInicial3 = mValorInicial3
End Property
Public Property Let ValorInicial3(ByVal Valor As Currency)
    mValorInicial3 = Valor
End Property


Public Property Get ValorFinal3() As Currency
    ValorFinal3 = mValorFinal3
End Property
Public Property Let ValorFinal3(ByVal Valor As Currency)
    mValorFinal3 = Valor
End Property


Public Property Get Percentual3() As Currency
    Percentual3 = mPercentual3
End Property
Public Property Let Percentual3(ByVal Valor As Currency)
    mPercentual3 = Valor
End Property


Public Property Get ValorInicial4() As Currency
    ValorInicial4 = mValorInicial4
End Property
Public Property Let ValorInicial4(ByVal Valor As Currency)
    mValorInicial4 = Valor
End Property


Public Property Get ValorFinal4() As Currency
    ValorFinal4 = mValorFinal4
End Property
Public Property Let ValorFinal4(ByVal Valor As Currency)
    mValorFinal4 = Valor
End Property


Public Property Get Percentual4() As Currency
    Percentual4 = mPercentual4
End Property
Public Property Let Percentual4(ByVal Valor As Currency)
    mPercentual4 = Valor
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
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nMesAno As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE [Mes Ano] = " & preparaTexto(nMesAno)
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
    xCondicao = " WHERE [Mes Ano] < " & preparaTexto(mMesAno)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Mes Ano] DESC")
    LocalizarAnterior = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE [Mes Ano] > " & preparaTexto("      ")
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
    xCondicao = " WHERE [Mes Ano] > " & preparaTexto(mMesAno)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Mes Ano] ASC")
    LocalizarProximo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE [Mes Ano] < " & preparaTexto("999999")
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Mes Ano] DESC")
    LocalizarUltimo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Tabela_Folha ( [Mes Ano], [Salario Minimo], [Cesta Basica], [Salario Familia 1], [Salario Familia Acima de], "
    gSQL = gSQL & "[Salario Familia 2], [Valor Inicial 1], [Valor Final 1], [Percentual 1], [Valor Inicial 2], "
    gSQL = gSQL & "[Valor Final 2], [Percentual 2], [Valor Inicial 3], [Valor Final 3], [Percentual 3], "
    gSQL = gSQL & "[Valor Inicial 4], [Valor Final 4], [Percentual 4], [Observacao 1], [Observacao 2], "
    gSQL = gSQL & "[Observacao 3] ) VALUES ( "
    Call sqlTexto(1, mMesAno, ", ")
    Call sqlValor(1, mSalarioMinimo, ", ")
    Call sqlValor(1, mCestaBasica, ", ")
    Call sqlValor(1, mSalarioFamilia1, ", ")
    Call sqlValor(1, mSalarioFamiliaAcimade, ", ")
    Call sqlValor(1, mSalarioFamilia2, ", ")
    Call sqlValor(1, mValorInicial1, ", ")
    Call sqlValor(1, mValorFinal1, ", ")
    Call sqlValor(1, mPercentual1, ", ")
    Call sqlValor(1, mValorInicial2, ", ")
    Call sqlValor(1, mValorFinal2, ", ")
    Call sqlValor(1, mPercentual2, ", ")
    Call sqlValor(1, mValorInicial3, ", ")
    Call sqlValor(1, mValorFinal3, ", ")
    Call sqlValor(1, mPercentual3, ", ")
    Call sqlValor(1, mValorInicial4, ", ")
    Call sqlValor(1, mValorFinal4, ", ")
    Call sqlValor(1, mPercentual4, ", ")
    Call sqlTexto(1, mObservacao1, ", ")
    Call sqlTexto(1, mObservacao2, ", ")
    Call sqlTexto(1, mObservacao3, " )")
    gConn.Execute gSQL, Incluir, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nMesAno As String) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Tabela_Folha SET "
    Call sqlTexto(2, "[Mes Ano] = ", mMesAno)
    Call sqlValor(2, ", [Salario Minimo] = ", mSalarioMinimo)
    Call sqlValor(2, ", [Cesta Basica] = ", mCestaBasica)
    Call sqlValor(2, ", [Salario Familia 1] = ", mSalarioFamilia1)
    Call sqlValor(2, ", [Salario Familia Acima de] = ", mSalarioFamiliaAcimade)
    Call sqlValor(2, ", [Salario Familia 2] = ", mSalarioFamilia2)
    Call sqlValor(2, ", [Valor Inicial 1] = ", mValorInicial1)
    Call sqlValor(2, ", [Valor Final 1] = ", mValorFinal1)
    Call sqlValor(2, ", [Percentual 1] = ", mPercentual1)
    Call sqlValor(2, ", [Valor Inicial 2] = ", mValorInicial2)
    Call sqlValor(2, ", [Valor Final 2] = ", mValorFinal2)
    Call sqlValor(2, ", [Percentual 2] = ", mPercentual2)
    Call sqlValor(2, ", [Valor Inicial 3] = ", mValorInicial3)
    Call sqlValor(2, ", [Valor Final 3] = ", mValorFinal3)
    Call sqlValor(2, ", [Percentual 3] = ", mPercentual3)
    Call sqlValor(2, ", [Valor Inicial 4] = ", mValorInicial4)
    Call sqlValor(2, ", [Valor Final 4] = ", mValorFinal4)
    Call sqlValor(2, ", [Percentual 4] = ", mPercentual4)
    Call sqlTexto(2, ", [Observacao 1] = ", mObservacao1)
    Call sqlTexto(2, ", [Observacao 2] = ", mObservacao2)
    Call sqlTexto(2, ", [Observacao 3] = ", mObservacao3)
    gSQL = gSQL & " WHERE [Mes Ano] = " & preparaTexto(nMesAno)
    gConn.Execute gSQL, Alterar, adCmdText + adExecuteNoRecords
    Alterar = True
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nMesAno As String) As Boolean

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Tabela_Folha"
    gSQL = gSQL & " WHERE [Mes Ano] = " & preparaTexto(nMesAno)
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT [Mes Ano], [Salario Minimo], [Cesta Basica], [Salario Familia 1], [Salario Familia Acima de], [Salario Familia 2], [Valor Inicial 1], [Valor Final 1], [Percentual 1], [Valor Inicial 2], [Valor Final 2], [Percentual 2], [Valor Inicial 3], [Valor Final 3], [Percentual 3], [Valor Inicial 4], [Valor Final 4], [Percentual 4], [Observacao 1], [Observacao 2], [Observacao 3] FROM Tabela_Folha" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsTabelaFolha = New adodb.Recordset
    rsTabelaFolha.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsTabelaFolha.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsTabelaFolha.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    MesAno = rsTabelaFolha("Mes Ano").Value
    SalarioMinimo = rsTabelaFolha("Salario Minimo").Value
    CestaBasica = rsTabelaFolha("Cesta Basica").Value
    SalarioFamilia1 = rsTabelaFolha("Salario Familia 1").Value
    SalarioFamiliaAcimade = rsTabelaFolha("Salario Familia Acima de").Value
    SalarioFamilia2 = rsTabelaFolha("Salario Familia 2").Value
    ValorInicial1 = rsTabelaFolha("Valor Inicial 1").Value
    ValorFinal1 = rsTabelaFolha("Valor Final 1").Value
    Percentual1 = rsTabelaFolha("Percentual 1").Value
    ValorInicial2 = rsTabelaFolha("Valor Inicial 2").Value
    ValorFinal2 = rsTabelaFolha("Valor Final 2").Value
    Percentual2 = rsTabelaFolha("Percentual 2").Value
    ValorInicial3 = rsTabelaFolha("Valor Inicial 3").Value
    ValorFinal3 = rsTabelaFolha("Valor Final 3").Value
    Percentual3 = rsTabelaFolha("Percentual 3").Value
    ValorInicial4 = rsTabelaFolha("Valor Inicial 4").Value
    ValorFinal4 = rsTabelaFolha("Valor Final 4").Value
    Percentual4 = rsTabelaFolha("Percentual 4").Value
    Observacao1 = rsTabelaFolha("Observacao 1").Value
    Observacao2 = rsTabelaFolha("Observacao 2").Value
    Observacao3 = rsTabelaFolha("Observacao 3").Value
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsTabelaFolha.Close
    Set rsTabelaFolha = Nothing
End Function

