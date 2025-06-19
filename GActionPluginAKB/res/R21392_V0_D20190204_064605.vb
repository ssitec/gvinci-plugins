Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R21392

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

    Public Shared Function R21392(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByRef msg As DataTable, ByRef P82022() As Object, ByRef P82023() As Object, ByRef P82024() As Object, ByRef P82027() As Object) As DataTable
    ' 
    ' 21392 - Valida Dias e % CF Risco Sac - varios pre-ped - Versão: 0
    ' 
        'On Error GoTo Err_R21392
        Dim sCurrComponent as String

        Dim C508086 As Object
        Dim C508086DataType As Short
        Dim C508087 As Object
        Dim C508087DataType As Short
        Dim C508088 As Object
        Dim C508088DataType As Short
        Dim C508089 As Boolean
        Dim C508089DataType As Short
        Dim C508090DataType() As Short
        Dim C508091 As Boolean
        Dim C508091DataType As Short
        Dim C508092 As Object
        Dim C508092DataType As Short
        Dim QueryC508093 As New Object
        Dim RsC508093 As New Object
        Dim C508093DataType() As Short
        Dim RsC508093_EOF As Boolean
        Dim C508094 As Boolean
        Dim C508094DataType As Short
        Dim C508095 As Boolean
        Dim C508095DataType As Short
        Dim C508096 As Short
        Dim C508096DataType As Short
        Dim C508096ReturnDataType() As Short

        Dim QueryC508097 As New Object
        Dim RsC508097 As New Object
        Dim C508097DataType() As Short
        Dim RsC508097_EOF As Boolean
        Dim C508098 As Object
        Dim C508098DataType As Short
        Dim C508099 As Object
        Dim C508099DataType As Short
        Dim C508100 As Object
        Dim C508100DataType As Short
        Dim C508101 As Object
        Dim C508101DataType As Short
        Dim P82022Current As Integer

        Dim P82023Current As Integer

        Dim P82024Current As Integer

        Dim P82027Current As Integer

        ReDim ReturnDatatype(0)

        GoTo Comp_C508086

    Comp_C508086:

        ' Ret
        sCurrComponent = "Ret"
        C508086 = 1
        C508086DataType = 4
        GoTo Comp_C508087

    Comp_C508087:

        ' PriVal
        sCurrComponent = "PriVal"
        C508087DataType = 4
        P82022Current = 1
        C508087 = True
        P82023Current = 1
        C508087 = True
        P82024Current = 1
        C508087 = True
        P82027Current = 1
        C508087 = True

        GoTo Comp_C508088

    Comp_C508088:

        ' FimVal
        sCurrComponent = "FimVal"
        C508088DataType = 4
        C508088 = (P82022Current > UBound(P82022))
        GoTo Comp_C508089

    Comp_C508089:

        ' FimVal=1?
        sCurrComponent = "FimVal=1?"
        C508089 = (fn_ConvertValueCompiled(C508088, C508088DataType, AKB_DecimalPoint, False) = 1)
        C508089DataType = AKBTypeConst.cAKBTypeLogical
        If C508089 Then
            GoTo Comp_C508090
        Else
            GoTo Comp_C508091
        End If

    Comp_C508090:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C508090 As DataTable = New DataTable()
        tb_C508090.Columns.Add("Ret" & "")
        Dim row_C508090 As DataRow = tb_C508090.NewRow()
        row_C508090(0) = C508086
        tb_C508090.Rows.Add(row_C508090)
        R21392 = tb_C508090
        ReDim C508090DataType(1)
        C508090DataType(1) = C508086DataType
        ReturnDataType = C508090DataType

        GoTo Exit_R21392

    Comp_C508091:

        ' Gerar=0?
        sCurrComponent = "Gerar=0?"
        C508091 = (fn_ConvertValueCompiled(P82022(P82022Current), 4, AKB_DecimalPoint, False) = 0)
        C508091DataType = AKBTypeConst.cAKBTypeLogical
        If C508091 Then
            GoTo Comp_C508092
        Else
            GoTo Comp_C508093
        End If

    Comp_C508092:

        ' ProVal
        sCurrComponent = "ProVal"
        C508092DataType = 4
        If Not (P82022Current > UBound(P82022)) Then
            P82022Current = P82022Current + 1
            C508092 = True
        Else
            C508092 = False
        End If
        If Not (P82023Current > UBound(P82023)) Then
            P82023Current = P82023Current + 1
            C508092 = True
        Else
            C508092 = False
        End If
        If Not (P82024Current > UBound(P82024)) Then
            P82024Current = P82024Current + 1
            C508092 = True
        Else
            C508092 = False
        End If
        If Not (P82027Current > UBound(P82027)) Then
            P82027Current = P82027Current + 1
            C508092 = True
        Else
            C508092 = False
        End If

        GoTo Comp_C508088

    Comp_C508093:

        ' FlagRS
        sCurrComponent = "FlagRS"
        QueryC508093 = con.CreateCommand()
        QueryC508093.CommandText = QueryC508093.CommandText & " " & "SELECT NVL(WF_CONDICAO_PGTO.E_Risco_Sacado,0) FROM AKBUSER01.WF_CONDICAO_PGTO WHERE WF_CONDICAO_PGTO.Cod_Pagto = " & _ 
ConvertValue(P82024(P82024Current), 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC508093.Transaction = txn
        RsC508093 = QueryC508093.ExecuteReader()
        Dim iC508093 As Short
        ReDim C508093DataType(RsC508093.FieldCount)
        For iC508093 = 0 to RsC508093.FieldCount - 1
            Select Case RsC508093.GetDataTypeName(iC508093).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C508093DataType(iC508093 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C508093DataType(iC508093 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C508093DataType(iC508093 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC508093
        RsC508093_EOF = Not RsC508093.Read()

        GoTo Comp_C508094

    Comp_C508094:

        ' FlagRS=0?
        sCurrComponent = "FlagRS=0?"
        C508094 = (fn_ConvertValueCompiled(RsC508093(0), C508093DataType(1), AKB_DecimalPoint, False) = 0)
        C508094DataType = AKBTypeConst.cAKBTypeLogical
        If C508094 Then
            GoTo Comp_C508092
        Else
            GoTo Comp_C508098
        End If

    Comp_C508095:

        ' DiasTaxa<>0?
        sCurrComponent = "DiasTaxa<>0?"
        C508095 = (fn_ConvertValueCompiled(C508099, C508099DataType, AKB_DecimalPoint, False) > 0 AND fn_ConvertValueCompiled(C508100, C508100DataType, AKB_DecimalPoint, False) > 0)
        C508095DataType = AKBTypeConst.cAKBTypeLogical
        If C508095 Then
            GoTo Comp_C508092
        Else
            GoTo Comp_C508096
        End If

    Comp_C508096:

        ' MSG1
        sCurrComponent = "MSG1"
        Dim row_C508096 As DataRow = msg.NewRow()
        row_C508096(0) = "Condição de pagamento " & _ 
P82024(P82024Current) & " , do pré-pedido " & _ 
P82023(P82023Current) & " é risco sacado, e não existe parametrização de taxa/dias de risco sacado para o cliente " & _ 
P82027(P82027Current) & " do pré-pedido."
        msg.Rows.Add(row_C508096)
        C508096DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C508101

    Comp_C508097:

        ' SelTaxasRS
        sCurrComponent = "SelTaxasRS"
        QueryC508097 = con.CreateCommand()
        QueryC508097.CommandText = QueryC508097.CommandText & " " & "SELECT NVL(Qtde_Dias_Vencimento,0), NVL(Custo_Financeiro_Mensal,0)"
        QueryC508097.CommandText = QueryC508097.CommandText & " " & "FROM WF_TAXA_RISCO_SACADO"
        QueryC508097.CommandText = QueryC508097.CommandText & " " & "WHERE Cod_Cliente = " & _ 
ConvertValue(P82027(P82027Current), 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC508097.CommandText = QueryC508097.CommandText & " " & "AND Data_Validade_Inicial <= " & _ 
ConvertValue(C508098, C508098DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC508097.CommandText = QueryC508097.CommandText & " " & "AND (Data_Validade_Final >= " & _ 
ConvertValue(C508098, C508098DataType, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " OR Data_Validade_Final IS NULL)"
        QueryC508097.Transaction = txn
        RsC508097 = QueryC508097.ExecuteReader()
        Dim iC508097 As Short
        ReDim C508097DataType(RsC508097.FieldCount)
        For iC508097 = 0 to RsC508097.FieldCount - 1
            Select Case RsC508097.GetDataTypeName(iC508097).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C508097DataType(iC508097 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C508097DataType(iC508097 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C508097DataType(iC508097 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC508097
        RsC508097_EOF = Not RsC508097.Read()

        GoTo Comp_C508099

    Comp_C508098:

        ' SysDate
        sCurrComponent = "SysDate"
        C508098DataType = 2
        C508098 = CDate(Format(GetDate(con, txn), "dd/MM/yyyy"))
        GoTo Comp_C508097

    Comp_C508099:

        ' QtdDiasRS
        sCurrComponent = "QtdDiasRS"
        C508099DataType = 0
        C508099 = RsC508097(0)
        C508099DataType = C508097Datatype(1)
        If C508099DataType = AKBTypeConst.cAKBTypeString Then
          C508099 = IIF(IsDBNull(C508099), C508099, Strings.RTrim(C508099))
        End If 

        GoTo Comp_C508100

    Comp_C508100:

        ' %CF_RS
        sCurrComponent = "%CF_RS"
        C508100DataType = 0
        C508100DataType = C508097Datatype(2)
        If C508100DataType = AKBTypeConst.cAKBTypeString And Not IsDBNull(RsC508097(1)) Then
          C508100 = Strings.RTrim(RsC508097(1))
        Else 
          C508100 = RsC508097(1)
        End If 

        GoTo Comp_C508095

    Comp_C508101:

        ' ATRIBUIVA1
        sCurrComponent = "ATRIBUIVA1"
        C508101DataType = 4
        C508086 = fn_ConvertValueCompiled(0, 4, AKB_DecimalPoint)
        C508101 = True
        GoTo Comp_C508090

    Exit_R21392:

        Exit Function

    Err_R21392:
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
                    fn_ConvertValueCompiled = IIf(CBool(vValue) = 1 Or CBool(vValue) = True, 1, 0)
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
    Public Shared Function ObjTable_NewIdentity(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByVal TableName as String) As Long
        Dim Rs As New Object
        Dim sSql As String
        Dim Query As New Object

        sSql = "Select SEQ_" & TableName & ".nextval from dual"
        Query = con.CreateCommand()
        Query.CommandText = sSql
        Query.Transaction = txn
        Rs = Query.ExecuteReader()

        ' Retorna a data
        If Rs.Read() Then
            ObjTable_NewIdentity = Rs(0)
        End If
        Rs.Close()
    End Function

End Class
