Attribute VB_Name = "ecf_mecaf"
Declare Function OpenCif Lib "ECF32M.DLL" () As Long
Declare Sub CloseCif Lib "ECF32M.DLL" ()
Declare Function AbreCupomFiscal Lib "ECF32M.DLL" () As Long
Declare Function LeituraX Lib "ECF32M.DLL" (ByVal RelGer As Byte) As Long
Declare Function ReducaoZ Lib "ECF32M.DLL" (ByVal RelGer As Byte) As Long
Declare Function ProgramaHorarioVerao Lib "ECF32M.DLL" (ByVal hv As Byte) As Long
Declare Function ProgramaLegenda Lib "ECF32M.DLL" (ByVal reg As String, ByVal leg As String) As Long
Declare Function TransStatus Lib "ECF32M.DLL" (ByVal BitTest As Long, ByVal BufStat As String) As Long
Declare Function TransTotCont Lib "ECF32M.DLL" () As Long
Declare Function TransDataHora Lib "ECF32M.DLL" () As Long
Declare Function TotalizarCupom Lib "ECF32M.DLL" (ByVal oper As Byte, ByVal toper As Byte, ByVal valor As String, ByVal legendaOp As String) As Long
Declare Function Pagamento Lib "ECF32M.DLL" (ByVal reg As String, ByVal vpgto As String, ByVal subtr As Byte) As Long
Declare Function LeMemFiscalData Lib "ECF32M.DLL" (ByVal datai As String, ByVal dataf As String, ByVal res As Byte) As Long
Declare Function ObtemRetorno Lib "ECF32M.DLL" (ByVal buf_ret As String) As Long
Declare Function VendaItem Lib "ECF32M.DLL" (ByVal fmt As Byte, ByVal qtd As String, ByVal punit As String, ByVal trib As String, ByVal TDesc As Byte, ByVal valor As String, ByVal unid As String, ByVal cod As String, ByVal ex As Byte, ByVal descr As String, ByVal legendaOp As String) As Long
Declare Function CancelamentoItem Lib "ECF32M.DLL" (ByVal numitem As String) As Long
Declare Function DescontoItem Lib "ECF32M.DLL" (ByVal toper As Byte, ByVal valor As String, ByVal legop As String) As Long
Declare Function FechaCupomFiscal Lib "ECF32M.DLL" (ByVal tam_msg As String, ByVal msg As String) As Long
Declare Function CancelaCupomFiscal Lib "ECF32M.DLL" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Function TrataRetorno(lngRet As Long) As String
    Dim strMsg As String
    Dim lngBufRet As Long
    Dim strBufRet As String
    
    strBufRet = ObtemRet(lngRet)
    
    strMsg = TraduzCodigoRetorno(lngRet)
'    MsgBox " " + Str(lngRet) + " - " + strMsg
    
    TrataRetorno = strBufRet
    
End Function
Public Function ObtemRet(lngRet As Long) As String
    Const MaxBuf = 2000
    Dim strBufRet As String * MaxBuf
    Dim intCont As Integer

    If lngRet = CIF_OK Then
        For intCont = 1 To 40
            lngRet = ObtemRetorno(strBufRet)
            If (lngRet <> -97) Then
                Exit For
            End If
            Sleep 1000
        Next intCont
    End If
    ObtemRet = strBufRet
End Function
Function TraduzCodigoRetorno(ByVal intretorno As Integer) As String
    Dim strMsg As String
    Select Case intretorno
    
        '---------------------------
        ' Codigo de retorno dos comandos da impressora
        '
        Case -1
            strMsg = "Cabe�alho cont�m caracteres inv�lidos"
        Case -2
            strMsg = "Comando inexistente"
        Case -3
            strMsg = "Valor n�o num�rico em campo num�rico"
        Case -4
            strMsg = "Valor fora da faixa entre 20h e 7Fh"
        Case -5
            strMsg = "Campo deve iniciar com @, & ou %"
        Case -6
            strMsg = "Campo deve iniciar com $, # ou ?"
        Case -7
            strMsg = "O intervalo � inconsistente. No caso de datas, valores anteriores a " & _
                   "01/01/95 ser�o consideradas como ano 2000 a 2094"
        Case -9
            strMsg = "A string TOTAL n�o � aceita"
        Case -10
            strMsg = "A sintaxe do comando est� errada"
        Case -11
            strMsg = "Excedeu n�mero m�ximo de linhas permitidas pelo comando"
        Case -12
            strMsg = "O terminador enviado n�o est� obedecendo o protocolo de comunica��o"
        Case -13
            strMsg = "O checksum est� incorreto"
        Case -15
            strMsg = "A situa��o tribut�ria deve iniciar com T, F, I ou N"
        Case -16
            strMsg = "Data inv�lida"
        Case -17
            strMsg = "Hora inv�lida"
        Case -18
            strMsg = "Al�quota n�o programada ou fora do intervalo"
        Case -19
            strMsg = "O campo de sinal est� incorreto"
        Case -20
            strMsg = "Comando s� aceito em Interven��o Fiscal"
        Case -22
            strMsg = "� necess�rio abrir o Cupom Fiscal"
        Case -23
            strMsg = "Comando n�o aceito durante Cupom Fiscal"
        Case -24
            strMsg = "� necess�rio abrir Cupom N�o Fiscal"
        Case -25
            strMsg = "Comando n�o aceito durante Cupom N�o Fiscal"
        Case -26
            strMsg = "O rel�gio j� est� em hor�rio de ver�o"
        Case -27
            strMsg = "O rel�gio n�o est� em hor�rio de ver�o"
        Case -28
            strMsg = "Necess�rio realizar Redu��o Z"
        Case -29
            strMsg = "Fechamento do dia (Redu��o Z) j� executado"
        Case -30
            strMsg = "Necess�rio programar legenda"
        Case -31
            strMsg = "Item inexistente ou j� cancelado"
        Case -32
            strMsg = "O cupom anterior n�o pode ser cancelado"
        Case -33
            strMsg = "Detectado falta de papel. Verifique a impressora."
        Case -36
            strMsg = "Necess�rio programar os dados do estabelecimento"
        Case -37
            strMsg = "Necess�rio realizar Interven��o Fiscal."
        Case -38
        strMsg = "Mem�ria Fiscal n�o permite mais realizar vendas. Apenas � poss�vel realizar LeituraX " & _
               "ou Leitura da Mem�ria Fiscal."
        Case -39
            strMsg = "Mem�ria Fiscal n�o permite mais realizar vendas. Apenas � poss�vel realizar LeituraX " & _
                   "ou Leitura da Mem�ria Fiscal, deve haver algum problema na mem�ria NOVRAM. Ser� " & _
                   "necess�rio realizar Interven��o Fiscal."
        Case -40
            strMsg = "Necess�rio programar a data do rel�gio"
        Case -41
            strMsg = "N�mero m�ximo de itens por cupom ultrapassado"
        Case -42
            strMsg = "J� foi realizado o Ajuste de Hora Di�rio"
        Case -43
            strMsg = "Comando v�lido ainda em execu��o"
        Case -44
            strMsg = "Est� em estado de Impress�o de Cheques"
        Case -45
            strMsg = "N�o est� em estado de Impress�o de Cheques"
        Case -46
            strMsg = "Necess�rio inserir o cheque"
        Case -47
            strMsg = "Necess�rio inserir nova bobina"
        Case -48
            strMsg = "Necess�rio executar uma Leitura X"
        Case -49
            strMsg = "Detectado algum problema na impressora (Paper jam, sobretens�o, etc)."
        Case -50
            strMsg = "Cupom j� totalizado"
        Case -51
            strMsg = "Necess�rio totalizar cupom antes de fechar"
        Case -52
            strMsg = "Necess�rio finalizar Cupom com comando correto"
        Case -53
            strMsg = "Ocorreu erro de grava��o na Mem�ria Fiscal"
        Case -54
            strMsg = "Excedeu n�mero m�ximo de estabelecimentos"
        Case -55
            strMsg = "Mem�ria fiscal n�o inicializada"
        Case -56
            strMsg = "Ultrapassou valor do pagamento"
        Case -57
            strMsg = "Registrador n�o programado ou troco j� realizado"
        Case -58
            strMsg = "Falta completar valor do pagamento"
        Case -59
            strMsg = "Campo somente de caracteres n�o num�ricos"
        Case -60
            strMsg = "Excedeu campo m�ximo de caracteres"
        Case -61
            strMsg = "Troco n�o realizado"
        Case -62
            strMsg = "Comando desabilitado"
            
        '---------------------------
        ' Codigo de retorno de funcoes da DLL
        '
        Case CIF_OK
            strMsg = "Opera��o efetuada com sucesso"
        Case CIF_PPAPEL
            strMsg = "Sucesso, detectado pouco papel"
        Case CIF_CANCCUP
            strMsg = "Sucesso, cancelando cupom"
        Case CIF_CUPNF
            strMsg = "Sucesso, abrindo cupom rel gerencial"
        Case CIF_ERR
            strMsg = "Falha geral na execu��o da DLL"
        Case CIF_EMEXECUCAO
            strMsg = "Comando v�lido ainda em execu��o"
        Case CIF_ERR_CONFIG
            strMsg = "Erro no arquivo CIF.INI"
        Case CIF_ERR_SERIAL
            strMsg = "Erro na abertura da serial"
        Case CIF_ERR_SYS
            strMsg = "Falha na aloca��o de recursos do Windows"
        Case CIF_ERR_ANSWER
            strMsg = "Retorno nao reconhecido"
        Case CIF_ERR_READSER
            strMsg = "Falha na leitura da serial"
        Case CIF_ERR_TEMP
            strMsg = "Temperatura da cabe�a de impress�o alta"
        Case CIF_ERR_PPAPEL
            strMsg = "Pouco papel"
        Case CIF_IRRECUPERAVEL
            strMsg = "Erro irrecuper�vel"
        Case CIF_ERR_MECANICO
            strMsg = "Erro mec�nico"
        Case CIF_ERR_TABERTA
            strMsg = "Tampa aberta"
        Case CIF_SEMRETORNO
            strMsg = "Opera��o sem retorno"
        Case CIF_OVERFLOW
            strMsg = "Buffer overflow. Tamanho da mensagem enviada pelo ECF � maior do que o buffer fornecido pela aplica��o"
        Case CIF_TIMEOUT
            strMsg = "TimeOut na execucao do comando"
        Case Else
            strMsg = "C�digo de retorno inexistente"
    End Select
    TraduzCodigoRetorno = strMsg
'    MsgBox strMsg
End Function
Function RetornaBStatus(ByVal NumBit As Long) As Boolean
    Dim lngRet As Long
    Dim strBuffer As String * 40
    Dim strBufferFormatado As String * 40
    Dim strByte1  As String * 8
    Dim strByte2  As String * 8
    Dim strByte3  As String * 8
    Dim strByte4  As String * 8
    Dim strByte5  As String * 8
    RetornaBStatus = False
    strBuffer = String(MaxSize, 0)
    lngRet = TransStatus(0, strBuffer)
    If lngRet <> CIF_OK Then
        Exit Function
    End If
    strByte1 = StrReverse(Mid(strBuffer, 1, 8))
    strByte2 = StrReverse(Mid(strBuffer, 9, 8))
    strByte3 = StrReverse(Mid(strBuffer, 17, 8))
    strByte4 = StrReverse(Mid(strBuffer, 25, 8))
    strByte5 = StrReverse(Mid(strBuffer, 33, 8))
    strBufferFormatado = strByte1 & strByte2 & strByte3 & strByte4 & strByte5
    If Mid$(strBufferFormatado, NumBit, 1) = "1" Then
        RetornaBStatus = True
    End If
End Function

