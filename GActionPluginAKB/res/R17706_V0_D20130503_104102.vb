Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R17706

    ' Tipo de dado do AKB
    Private Enum AKBTypeConst
        cAKBTypeNone = -1
        cAKBTypeNumeric = 1
        cAKBTypeDate = 2
        cAKBTypeString = 3
        cAKBTypeLogical = 4

        ' atv 13.453-107.368
        cAKBTypeUndefined = 0
        cAKBTypeQuery = 5
    End Enum

    Public Const AKB_DecimalPoint = ","
    Public Const DecimalSymbol = "."
    Public Const DateIdentifier = "'"
    Public Const StringIdentifier = "'"
    Public Shared Sub Main ()

    End Sub

    Public Shared Function R17706(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal P64333 As Double, ByVal P64334 As Double) As DataTable
    ' 
    ' 17706 - Gravar Codigo Produto NF-e Pedido Itens - Versão: 0
    ' 
        'On Error GoTo Err_R17706
        Dim sCurrComponent as String

        Dim C392951 As Object
        Dim C392951DataType As Short
        Dim QueryC392952 As New Object
        Dim RsC392952 As New Object
        Dim C392952DataType() As Short
        Dim RsC392952_EOF As Boolean
        Dim C392953 As Object
        Dim C392953DataType As Short
        Dim C392954 As Boolean
        Dim C392954DataType As Short
        Dim C392955 As Object
        Dim C392955DataType As Short
        Dim C392957 As Object
        Dim C392957DataType As Short
        Dim QueryC392958 As New Object
        Dim C392958 As Integer
        Dim C392958DataType As Short
        Dim C392960DataType() As Short
        ReDim ReturnDatatype(0)

        GoTo Comp_C392951

    Comp_C392951:

        ' vRet
        sCurrComponent = "vRet"
        C392951 = 1
        C392951DataType = 4
        GoTo Comp_C392953

    Comp_C392952:

        ' selDesconto280
        sCurrComponent = "selDesconto280"
        QueryC392952 = con.CreateCommand()
        QueryC392952.CommandText = QueryC392952.CommandText & " " & "SELECT COUNT(*)"
        QueryC392952.CommandText = QueryC392952.CommandText & " " & "FROM AKBUSER01.WF_PED_SEQ_DESCONTO"
        QueryC392952.CommandText = QueryC392952.CommandText & " " & "WHERE WF_PED_SEQ_DESCONTO.Tp_Produto = " & _ 
ConvertValue(P64333, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC392952.CommandText = QueryC392952.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.Pedido = " & _ 
ConvertValue(P64334, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC392952.CommandText = QueryC392952.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.Seq = 1"
        QueryC392952.CommandText = QueryC392952.CommandText & " " & "AND WF_PED_SEQ_DESCONTO.Tipo_Desconto = 280"
        QueryC392952.Transaction = txn
        RsC392952 = QueryC392952.ExecuteReader()
        Dim iC392952 As Short
        ReDim C392952DataType(RsC392952.FieldCount)
        For iC392952 = 0 to RsC392952.FieldCount - 1
            Select Case RsC392952.GetDataTypeName(iC392952).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C392952DataType(iC392952 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C392952DataType(iC392952 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C392952DataType(iC392952 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC392952
        RsC392952_EOF = Not RsC392952.Read()

        GoTo Comp_C392954

    Comp_C392953:

        ' qryCampoCodNF-e
        sCurrComponent = "qryCampoCodNF-e"
        C392953 = System.DBNull.Value
        C392953DataType = 5
        GoTo Comp_C392952

    Comp_C392954:

        ' Tem Desconto 280?
        sCurrComponent = "Tem Desconto 280?"
        C392954 = (fn_ConvertValueCompiled(RsC392952(0), C392952DataType(1), AKB_DecimalPoint, False) > 0)
        C392954DataType = AKBTypeConst.cAKBTypeLogical
        If C392954 Then
            GoTo Comp_C392955
        Else
            GoTo Comp_C392957
        End If

    Comp_C392955:

        ' atXprod2
        sCurrComponent = "atXprod2"
        C392955DataType = 4
        C392953 = fn_ConvertValueCompiled("WF_EMB_COMP_VDA_PROD.cProd2", 5, AKB_DecimalPoint)
        C392955 = True
        GoTo Comp_C392958

    Comp_C392957:

        ' atXprod
        sCurrComponent = "atXprod"
        C392957DataType = 4
        C392953 = fn_ConvertValueCompiled("WF_EMB_COMP_VDA_PROD.cProd", 5, AKB_DecimalPoint)
        C392957 = True
        GoTo Comp_C392958

    Comp_C392958:

        ' upCodNF-e PedidoItens
        sCurrComponent = "upCodNF-e PedidoItens"
        QueryC392958 = con.CreateCommand()
        QueryC392958.CommandText = QueryC392958.CommandText & " " & "UPDATE AKBUSER01.WF_PEDIDO_ITENS"
        QueryC392958.CommandText = QueryC392958.CommandText & " " & "SET WF_PEDIDO_ITENS.cProd = "
        QueryC392958.CommandText = QueryC392958.CommandText & " " & "(SELECT " & _ 
ConvertValue(C392953, C392953DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC392958.CommandText = QueryC392958.CommandText & " " & "FROM AKBUSER01.WF_EMB_COMP_VDA_PROD"
        QueryC392958.CommandText = QueryC392958.CommandText & " " & "WHERE WF_EMB_COMP_VDA_PROD.Sigla_Prod = WF_PEDIDO_ITENS.Sigla_Prod"
        QueryC392958.CommandText = QueryC392958.CommandText & " " & "AND WF_EMB_COMP_VDA_PROD.Acesso = WF_PEDIDO_ITENS.Acesso"
        QueryC392958.CommandText = QueryC392958.CommandText & " " & "AND WF_EMB_COMP_VDA_PROD.Cod_Embalagem = WF_PEDIDO_ITENS.Cod_Embalagem)"
        QueryC392958.CommandText = QueryC392958.CommandText & " " & "WHERE WF_PEDIDO_ITENS.Tp_Produto = " & _ 
ConvertValue(P64333, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC392958.CommandText = QueryC392958.CommandText & " " & "AND WF_PEDIDO_ITENS.Pedido = " & _ 
ConvertValue(P64334, 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & ""
        QueryC392958.Transaction = txn
        C392958 = QueryC392958.ExecuteNonQuery()
        C392958DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C392960

    Comp_C392960:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C392960 As DataTable = New DataTable()
        tb_C392960.Columns.Add("vRet" & "")
        Dim row_C392960 As DataRow = tb_C392960.NewRow()
        row_C392960(0) = C392951
        tb_C392960.Rows.Add(row_C392960)
        R17706 = tb_C392960
        ReDim C392960DataType(1)
        C392960DataType(1) = C392951DataType
        ReturnDataType = C392960DataType

        GoTo Exit_R17706

    Exit_R17706:

        Exit Function

    Err_R17706:
        If Not IsNothing(txn) Then
            txn.Rollback()
        End If
    End Function
    '
    ' Converte um valor String para o tipo apropriado
    '
    Public Shared Function fn_ConvertValueCompiled(ByVal vValue As Object, ByVal lDatatype As Integer, ByVal sDecimalPoint As String, _
                                            Optional bExpression As Boolean = false) As Object

        If Not IsDBNull(vValue) Then

            ' Verifica o tipo do dado
            Select Case lDatatype
                Case AKBTypeConst.cAKBTypeNumeric
                    If CStr(vValue) <> "" Then
                        fn_ConvertValueCompiled = CDbl(vValue)
                    End If
                Case AKBTypeConst.cAKBTypeDate
                    fn_ConvertValueCompiled = CDate(vValue)
                Case AKBTypeConst.cAKBTypeString, AKBTypeConst.cAKBTypeQuery
                    fn_ConvertValueCompiled = vValue
                Case AKBTypeConst.cAKBTypeLogical
                    fn_ConvertValueCompiled = IIf(vValue = 1 Or vValue = True, 1, 0)
                Case Else
                    If CStr(vValue) = "True" Or CStr(vValue) = "False" Then
                        fn_ConvertValueCompiled = IIf(vValue = 1 Or vValue = True, 1, 0)
                    ElseIf IsNumeric(vValue) Then
                        fn_ConvertValueCompiled = CDbl(vValue)
                    ElseIf IsDate(vValue) Then
                        fn_ConvertValueCompiled = CDate(vValue)
                    Else
                        fn_ConvertValueCompiled = vValue
                    End If
            End Select
        Else
            If Not bExpression Then
                fn_ConvertValueCompiled = System.DBNull.Value
            Else

                ' Verifica o tipo do dado
                Select Case lDatatype
                    Case AKBTypeConst.cAKBTypeNumeric
                        fn_ConvertValueCompiled = 0
                    Case AKBTypeConst.cAKBTypeDate
                        fn_ConvertValueCompiled = ""
                    Case AKBTypeConst.cAKBTypeString, AKBTypeConst.cAKBTypeQuery
                        fn_ConvertValueCompiled = ""
                    Case AKBTypeConst.cAKBTypeLogical
                        fn_ConvertValueCompiled = 0
                    Case Else
                        fn_ConvertValueCompiled = ""
                End Select
            End If
        End If
    End Function
    '
    ' Retorna o inteiro maior ou igual ao número
    '
    Public Shared Function fn_Ceiling(ByVal dParam As Double) As Integer

        ' Obtém valor
        If dParam <= 0 Then
            fn_Ceiling = Fix(dParam)
        Else
            If Int(dParam) = dParam Then
                fn_Ceiling = dParam
            Else
                fn_Ceiling = Fix(dParam) + 1
            End If
        End If
    End Function
    '
    ' Retorna o inteiro menor ou igual ao número
    '
    Public Shared Function fn_Floor(ByVal dParam As Double) As Integer

        ' Obtém valor
        If dParam >= 0 Then
            fn_Floor = Fix(dParam)
        Else
            If Int(dParam) = dParam Then
                fn_Floor = dParam
            Else
                fn_Floor = Fix(dParam) - 1
            End If
        End If
    End Function
    Public Shared ReturnDatatype() As Short
    Private Shared iColumns As Short
    Private Shared iCol As Short
    Public Shared Function GetReturnDatatype() As Short()
        GetReturnDatatype = ReturnDatatype
    End Function

    '
    'Retorna o valor por extenso
    '
    Public Shared Function Extenso(valor)

       Dim g(10, 1) As String
       Dim CEN(9) As String
       Dim UNI(9) As String
       Dim DEZ(19) As String
       Dim PARTE(9) As String
       Dim vlr As Double
       Dim Numero$
       Dim p
       Dim i
       Dim TEXTO
       Dim c, D, U
       Dim DE
       Dim VALOR_ANT
       Dim VALOR_POS
       Dim X1, X2, X3
       Dim E1, E2, E3
       Dim GRANDEZA

       If Not IsNumeric(valor) Then
           Extenso = "Valor Não Numerico"
           Exit Function
       End If

       g(1, 0) = "CENTAVO"
       g(2, 0) = ""
       g(3, 0) = "MIL"
       g(4, 0) = "MILHÃO"
       g(5, 0) = "BILHÃO"
       g(6, 0) = "TRILHÃO"

       g(1, 1) = "CENTAVOS"
       g(2, 1) = ""
       g(3, 1) = "MIL"
       g(4, 1) = "MILHÕES"
       g(5, 1) = "BILHÕES"
       g(6, 1) = "TRILHÕES"

       CEN(0) = ""
       CEN(1) = "CENTO"
       CEN(2) = "DUZENTOS"
       CEN(3) = "TREZENTOS"
       CEN(4) = "QUATROCENTOS"
       CEN(5) = "QUINHENTOS"
       CEN(6) = "SEISCENTOS"
       CEN(7) = "SETECENTOS"
       CEN(8) = "OITOCENTOS"
       CEN(9) = "NOVECENTOS"

       DEZ(0) = ""
       DEZ(1) = "DEZ"
       DEZ(2) = "VINTE"
       DEZ(3) = "TRINTA"
       DEZ(4) = "QUARENTA"
       DEZ(5) = "CINQUENTA"
       DEZ(6) = "SESSENTA"
       DEZ(7) = "SETENTA"
       DEZ(8) = "OITENTA"
       DEZ(9) = "NOVENTA"
       DEZ(11) = "ONZE"
       DEZ(12) = "DOZE"
       DEZ(13) = "TREZE"
       DEZ(14) = "QUATORZE"
       DEZ(15) = "QUINZE"
       DEZ(16) = "DEZESSEIS"
       DEZ(17) = "DEZESSETE"
       DEZ(18) = "DEZOITO"
       DEZ(19) = "DEZENOVE"

       UNI(0) = ""
       UNI(1) = "UM"
       UNI(2) = "DOIS"
       UNI(3) = "TRES"
       UNI(4) = "QUATRO"
       UNI(5) = "CINCO"
       UNI(6) = "SEIS"
       UNI(7) = "SETE"
       UNI(8) = "OITO"
       UNI(9) = "NOVE"

       vlr = CDbl(valor)
       Numero$ = Format$(valor, "000000000000000.00")
       p = 1
       For i = 6 To 2 Step -1
          PARTE$(i) = Mid$(Numero$, p, 3)
          p = p + 3
       Next
       PARTE$(1) = 0 & Right$(Numero$, 2)

       TEXTO = ""

       For i = 6 To 1 Step -1
           If Val(PARTE$(i)) <> 0 Then
              c = Val(Left$(PARTE$(i), 1))
              D = Val(Mid$(PARTE$(i), 2, 1))
              U = Val(Mid$(PARTE$(i), 3, 1))
              DE = Val(Mid$(PARTE$(i), 2, 2))
              VALOR_ANT = "0"
              VALOR_POS = "0"
              If i < 6 Then
                 VALOR_ANT = Left$(Numero$, (6 - i) * 3)
              End If
              If i > 2 Then
                 VALOR_POS = Mid$(Numero$, (7 - i) * 3 + 1, (i - 2) * 3)
              End If
              X1 = ""
              X2 = ""
              X3 = ""
              If c = 1 And DE = 0 Then
                   X1 = "CEM"
              Else
                   X1 = CEN(c)
              End If
              If DE < 11 Or DE > 19 Then
                   X2 = DEZ(D)
                   X3 = UNI(U)
              Else
                   X2 = DEZ(DE)
              End If
              'DETERMINA E's
              E1 = " "
              E2 = ""
              E3 = ""
              'Entre as centenas, dezenas e unidades
              If c > 0 And DE > 0 Then
                   E2 = " E "
              End If
              If D > 1 And U > 0 Then
                   E3 = " E "
              End If

              'Entre grandezas diferentes
              If i > 1 And c = 0 And DE <> 0 And Val(VALOR_ANT) <> 0 And Val(VALOR_POS) = 0 Then
                   E2 = " E "
              End If
              If i > 1 And c > 0 And DE = 0 And Val(VALOR_ANT) <> 0 And Val(VALOR_POS) = 0 Then
                   E1 = " E "
              End If

              'Entre a parte inteira e os centavos
              If i = 1 And DE <> 0 And Int(valor) <> 0 Then
                   E1 = " E "
              End If
              If Val(PARTE$(i)) > 1 Then
                   GRANDEZA = g(i, 1)
              Else
                   GRANDEZA = g(i, 0)
              End If
              TEXTO = TEXTO & E1 & X1 & E2 & X2 & E3 & X3 & " " & GRANDEZA
           End If
           If i = 2 Then
               If Int(valor) = 1 Then
                   TEXTO = TEXTO & " REAL"
               Else
                   If PARTE$(3) = "000" And PARTE$(2) = "000" Then
                       TEXTO = TEXTO & " DE REAIS"
                   Else
                       TEXTO = TEXTO & " REAIS"
                   End If
               End If
           End If

       Next
       Extenso = Trim(TEXTO)
    End Function
    '
    ' Validar CGC
    '
    Public Shared Function CalculaCGC(Numero As String) As String

    Dim i As Short
    Dim prod As Short
    Dim mult As Short
    Dim digito As Short

    If Not IsNumeric(Numero) Then
      CalculaCGC = """"
      Exit Function
    End If

    mult = 2
    For i = Len(Numero) To 1 Step -1
     prod = prod + Val(Mid(Numero, i, 1)) * mult
     mult = IIf(mult = 9, 2, mult + 1)
    Next

    digito = 11 - Int(prod Mod 11)
    digito = IIf(digito = 10 Or digito = 11, 0, digito)

    CalculaCGC = Trim(Str(digito))

    End Function
    '
    'Validar CGC
    '
    Public Shared Function ValidaCGC(CGC As String) As Boolean

       'Retira os caracteres de formatação
       CGC = Replace(CGC, ".", "")
       CGC = Replace(CGC, "/", "")
       CGC = Replace(CGC, "-", "")

    If CalculaCGC(Left(CGC, 12)) <> Mid(CGC, 13, 1) Then
      ValidaCGC = False
      Exit Function
    End If

    If CalculaCGC(Left(CGC, 13)) <> Mid(CGC, 14, 1) Then
      ValidaCGC = False
      Exit Function
    End If

    ValidaCGC = True

    End Function

    '
    'Validar CPF
    '
    Public Shared Function calculacpf(CPF As String) As Boolean

       'Retira os caracteres de formatação
       CPF = Replace(CPF, ".", "")
       CPF = Replace(CPF, "/", "")
       CPF = Replace(CPF, "-", "")

       'Esta rotina foi adaptada da revista Fórum Access
       On Error GoTo Err_CPF
       Dim i As Short        'utilizada nos FOR... NEXT
       Dim strcampo As String  'armazena do CPF que será utilizada para o cálculo
       Dim strCaracter As String   'armazena os dígitos do CPF da direita para a esquerda
       Dim intNumero As Short    'armazena o digito separado para cálculo (uma a um)
       Dim intMais As Short  'armazena o digito específico multiplicado pela sua base
       Dim lngSoma As Integer     'armazena a soma dos dígitos multiplicados pela sua base(intmais)
       Dim dblDivisao As Double    'armazena a divisão dos dígitos * base por 11
       Dim lngInteiro As Integer  'armazena inteiro da divisão
       Dim intResto As Short     'armazena o resto
       Dim intDig1 As Short  'armazena o 1º digito verificador
       Dim intDig2 As Short  'armazena o 2º digito verificador
       Dim strConf As String   'armazena o digito verificador

       lngSoma = 0
       intNumero = 0
       intMais = 0
       strcampo = Left(CPF, 9)

       'Inicia cálculos do 1º dígito
       For i = 2 To 10
           strCaracter = Right(strcampo, i - 1)
           intNumero = Left(strCaracter, 1)
           intMais = intNumero * i
           lngSoma = lngSoma + intMais
       Next i
       dblDivisao = lngSoma / 11

       lngInteiro = Int(dblDivisao) * 11
       intResto = lngSoma - lngInteiro
       If intResto = 0 Or intResto = 1 Then
           intDig1 = 0
       Else
           intDig1 = 11 - intResto
       End If

       strcampo = strcampo & intDig1 'concatena o CPF com o primeiro digito verificador
       lngSoma = 0
       intNumero = 0
       intMais = 0
       'Inicia cálculos do 2º dígito
       For i = 2 To 11
           strCaracter = Right(strcampo, i - 1)
           intNumero = Left(strCaracter, 1)
           intMais = intNumero * i
           lngSoma = lngSoma + intMais
       Next i
       dblDivisao = lngSoma / 11
       lngInteiro = Int(dblDivisao) * 11
       intResto = lngSoma - lngInteiro
       If intResto = 0 Or intResto = 1 Then
           intDig2 = 0
       Else
           intDig2 = 11 - intResto
       End If
       strConf = intDig1 & intDig2
       'Caso o CPF esteja errado dispara a mensagem
       If strConf <> Right(CPF, 2) Then
           calculacpf = False
       Else
           calculacpf = True
       End If
       Exit Function

    Exit_CPF:
       Exit Function
    Err_CPF:
       MsgBox(Err.Description)
       Resume Exit_CPF
    End Function
    '
    'Converte um valor para uma String no formato esperado pelo Banco
    '
    Public Shared Function ConvertValue(ByVal Value As Object, ByVal ValueType As Integer, ByVal DecimalSymbol As String, ByVal DateIdentifier As String, ByVal StringIdentifier As String, ByVal DateFormat as String) As String

       On Error GoTo ConvertValue_Error

       Dim sTemp As String

       If Not IsDBNull(Value) Then

           'Verifica o tipo do dado
           Select Case ValueType

               Case AKBTypeConst.cAKBTypeNumeric

                   'Substitui o ponto pelo separador do banco
                   sTemp = Trim(Str(Value))
                   ConvertValue = Replace(sTemp, ".", DecimalSymbol)

               Case AKBTypeConst.cAKBTypeDate
                   ConvertValue = "TO_DATE(" & DateIdentifier & Format(CDate(Value), DateFormat) & DateIdentifier & ",'dd-mm-yyyy hh24:mi:ss')"

               Case AKBTypeConst.cAKBTypeString
                   ConvertValue = StringIdentifier & Replace(Value, StringIdentifier, StrDup(2, StringIdentifier)) & StringIdentifier

               Case AKBTypeConst.cAKBTypeLogical
                   ConvertValue = IIf(Value, 1, 0)

               Case Else
                   ConvertValue = Value

           End Select
       Else

           'Se for nulo retorna a String "Null"
           ConvertValue = "Null"
       End If

       Exit Function

    ConvertValue_Error:
       MsgBox("Erro ao executar a versão compilada GVINCI da função ConvertValue ")
    End Function
    '
    ' Retorna a data atual do banco de dados
    '
    Public Shared Function GetDate(ByRef con As OracleConnection, ByRef txn As OracleTransaction) As Object
        Dim Rs As New Object
        Dim sSql As String
        Dim Query As New Object

        sSql = "Select sysdate from dual"
        Query = con.CreateCommand()
        Query.CommandText = sSql
        Query.Transaction = txn
        Rs = Query.ExecuteReader()

        ' Retorna a data
        If Rs.Read() Then
            GetDate = Rs(0)
        End If
        Rs.Close()
    End Function
    '
    ' Retorna a data atual do banco de dados
    '
    Public Shared Function ObjTable_NewIdentity(ByRef con As OracleConnection, ByVal TableName as String) As Long
        Dim Rs As New Object
        Dim sSql As String
        Dim Query As New Object

        sSql = "Select SEQ_" & TableName & ".nextval from dual"
        Query = con.CreateCommand()
        Query.CommandText = sSql
        Rs = Query.ExecuteReader()

        ' Retorna a data
        If Rs.Read() Then
            ObjTable_NewIdentity = Rs(0)
        End If
        Rs.Close()
    End Function

End Class
