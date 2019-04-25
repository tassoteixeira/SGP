Attribute VB_Name = "BalancaFilizola"
'Declaração das funções (tem que ser no módulo)

Declare Function ConfiguraBalanca Lib "PcScale.dll" (ByVal balanca As Integer, ByVal Handle As Long) As Boolean
Declare Function InicializaLeitura Lib "PcScale.dll" (ByVal balanca As Integer) As Boolean
Declare Function ObtemInformacao Lib "PcScale.dll" (ByVal balanca As Integer, ByVal campo As Integer) As Double
Declare Function FinalizaLeitura Lib "PcScale.dll" (ByVal balanca As Integer) As Boolean
Declare Function EnviaPrecoCS Lib "PcScale.dll" (ByVal balanca As Integer, ByVal preco As Double) As Boolean
Declare Sub ExibeMsgErro Lib "PcScale.dll" (ByVal Handle As Long)

'As funções abaixo não são necessárias. Elas estão sendo usadas
'somente para exiber a configuração da balança
Declare Sub ObtemNomeBalanca Lib "PcScale.dll" (ByVal Modelo As Integer, ByVal Nome As String)
Declare Function ObtemParametrosBalanca Lib "PcScale" (ByVal balanca As Integer, ByRef Modelo As Integer, ByRef Porta As Integer, ByRef BaudRate As Long) As Boolean


