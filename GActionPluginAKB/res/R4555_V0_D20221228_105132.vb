Imports Microsoft.VisualBasic
Imports System.Data.OracleClient
Imports Oracle.DataAccess.Client
Imports System
Imports System.Data

Public Class clsRuleDynamicallyCompiled_R4555

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

    Public Shared Function R4555(ByRef con As OracleConnection, ByRef txn As OracleTransaction, ByRef msg As DataTable, ByRef P12595() As Object, ByRef P28838() As Object, ByRef P28839() As Object, ByRef P68165() As Object, ByVal P76805 As Double, ByRef P89540() As Object, ByRef P89541() As Object) As DataTable
    ' 
    ' 4555 - Gerar Pedido do Pré Pedido - Múltiplos - Versão: 0
    ' 
        'On Error GoTo Err_R4555
        Dim sCurrComponent as String

        Dim C54079 As Object
        Dim C54079DataType As Short
        Dim C54080 As Object
        Dim C54080DataType As Short
        Dim C54081 As Object
        Dim C54081DataType As Short
        Dim C54082 As Boolean
        Dim C54082DataType As Short
        Dim C54083DataType() As Short
        Dim C54091 As DataTable
        Dim C54091CurrentRow As Integer
        Dim C54091DataType() As Short
        Dim C103463 As Boolean
        Dim C103463DataType As Short
        Dim QueryC164205 As New Object
        Dim RsC164205 As New Object
        Dim C164205DataType() As Short
        Dim RsC164205_EOF As Boolean
        Dim C164206 As Object
        Dim C164206DataType As Short
        Dim C164207 As Boolean
        Dim C164207DataType As Short
        Dim C164208 As Object
        Dim C164208DataType As Short
        Dim C164209 As Double
        Dim C164209DataType As Short
        Dim C164210 As Object
        Dim C164210DataType As Short
        Dim C164211 As Boolean
        Dim C164211DataType As Short
        Dim C164212 As Short
        Dim C164212DataType As Short
        Dim C164212ReturnDataType() As Short

        Dim C164213 As Object
        Dim C164213DataType As Short
        Dim C164214 As Double
        Dim C164214DataType As Short
        Dim C164215 As Object
        Dim C164215DataType As Short
        Dim C164216 As Boolean
        Dim C164216DataType As Short
        Dim P12595Current As Integer

        Dim P28838Current As Integer

        Dim P28839Current As Integer

        Dim P68165Current As Integer

        Dim P89540Current As Integer

        Dim P89541Current As Integer

        ReDim ReturnDatatype(0)

        GoTo Comp_C164213

    Comp_C54079:

        ' First
        sCurrComponent = "First"
        C54079DataType = 4
        P12595Current = 1
        C54079 = True
        P28838Current = 1
        C54079 = True
        P28839Current = 1
        C54079 = True
        P68165Current = 1
        C54079 = True
        P89541Current = 1
        C54079 = True
        P89540Current = 1
        C54079 = True

        GoTo Comp_C54080

    Comp_C54080:

        ' Eof
        sCurrComponent = "Eof"
        C54080DataType = 4
        C54080 = (P12595Current > UBound(P12595))
        GoTo Comp_C54082

    Comp_C54081:

        ' Next
        sCurrComponent = "Next"
        C54081DataType = 4
        If Not (P12595Current > UBound(P12595)) Then
            P12595Current = P12595Current + 1
            C54081 = True
        Else
            C54081 = False
        End If
        If Not (P28838Current > UBound(P28838)) Then
            P28838Current = P28838Current + 1
            C54081 = True
        Else
            C54081 = False
        End If
        If Not (P28839Current > UBound(P28839)) Then
            P28839Current = P28839Current + 1
            C54081 = True
        Else
            C54081 = False
        End If
        If Not (P68165Current > UBound(P68165)) Then
            P68165Current = P68165Current + 1
            C54081 = True
        Else
            C54081 = False
        End If
        If Not (P89541Current > UBound(P89541)) Then
            P89541Current = P89541Current + 1
            C54081 = True
        Else
            C54081 = False
        End If
        If Not (P89540Current > UBound(P89540)) Then
            P89540Current = P89540Current + 1
            C54081 = True
        Else
            C54081 = False
        End If

        GoTo Comp_C54080

    Comp_C54082:

        ' Eof=1
        sCurrComponent = "Eof=1"
        C54082 = (fn_ConvertValueCompiled(C54080, C54080DataType, AKB_DecimalPoint, False) = 1)
        C54082DataType = AKBTypeConst.cAKBTypeLogical
        If C54082 Then
            GoTo Comp_C164211
        Else
            GoTo Comp_C103463
        End If

    Comp_C54083:

        ' RET1
        sCurrComponent = "RET1"
        Dim tb_C54083 As DataTable = New DataTable()
        tb_C54083.Columns.Add("Eof=1" & "")
        Dim row_C54083 As DataRow = tb_C54083.NewRow()
        row_C54083(0) = C54082
        tb_C54083.Rows.Add(row_C54083)
        R4555 = tb_C54083
        ReDim C54083DataType(1)
        C54083DataType(1) = C54082DataType
        ReturnDataType = C54083DataType

        GoTo Exit_R4555

    Comp_C54091:

        ' GeraPedido
        sCurrComponent = "GeraPedido"
        C54091 = clsRuleDynamicallyCompiled_R4557.R4557(con, txn, msg, IIf(Not IsDBNull(P12595(P12595Current)), P12595(P12595Current), System.DBNull.Value), System.DBNull.Value, System.DBNull.Value, fn_ConvertValueCompiled( "0", 4, AKB_DecimalPoint), IIf(Not IsDBNull(P28839(P28839Current)), P28839(P28839Current), System.DBNull.Value), System.DBNull.Value, IIf(Not IsDBNull(P68165(P68165Current)), P68165(P68165Current), System.DBNull.Value), IIf(Not IsDBNull(P76805), P76805, System.DBNull.Value), System.DBNull.Value, IIf(Not IsDBNull(P89541(P89541Current)), P89541(P89541Current), System.DBNull.Value), IIf(Not IsDBNull(P89540(P89540Current)), P89540(P89540Current), System.DBNull.Value))
        C54091CurrentRow = 1
        iColumns = 0
        If Not IsNothing(C54091) Then
          iColumns = C54091.Columns.Count
        End If
        ReDim C54091DataType(iColumns)
        For iCol = 1 To iColumns
            C54091DataType(iCol) = clsRuleDynamicallyCompiled_R4557.GetReturnDatatype(iCol)
        Next iCol

        GoTo Comp_C164214

    Comp_C103463:

        ' Gerar?
        sCurrComponent = "Gerar?"
        C103463 = (( fn_ConvertValueCompiled(P28838(P28838Current), 4, AKB_DecimalPoint, False) = 1 ))
        C103463DataType = AKBTypeConst.cAKBTypeLogical
        If C103463 Then
            GoTo Comp_C164205
        Else
            GoTo Comp_C54081
        End If

    Comp_C164205:

        ' Codigo Moeda IS NOT NULL
        sCurrComponent = "Codigo Moeda IS NOT NULL"
        QueryC164205 = con.CreateCommand()
        QueryC164205.CommandText = QueryC164205.CommandText & " " & "SELECT WF_PRE_PEDIDO.Codigo_Moeda "
        QueryC164205.CommandText = QueryC164205.CommandText & " " & "FROM AKBUSER01.WF_PRE_PEDIDO "
        QueryC164205.CommandText = QueryC164205.CommandText & " " & "WHERE"
        QueryC164205.CommandText = QueryC164205.CommandText & " " & "WF_PRE_PEDIDO.Id_PrePedido = " & _ 
ConvertValue(P12595(P12595Current), 1, DecimalSymbol, DateIdentifier, StringIdentifier, "dd-MM-yyyy HH:mm:ss") & " "
        QueryC164205.CommandText = QueryC164205.CommandText & " " & "AND WF_PRE_PEDIDO.Codigo_Moeda IS NOT NULL"
        QueryC164205.Transaction = txn
        RsC164205 = QueryC164205.ExecuteReader()
        Dim iC164205 As Short
        ReDim C164205DataType(RsC164205.FieldCount)
        For iC164205 = 0 to RsC164205.FieldCount - 1
            Select Case RsC164205.GetDataTypeName(iC164205).ToUpper()
                Case "NUMBER", "INT32", "DECIMAL", "INT16", "DOUBLE"
                    C164205DataType(iC164205 + 1) = AKBTypeConst.cAKBTypeNumeric
                Case "DATE", "TIMESTAMP"
                    C164205DataType(iC164205 + 1) = AKBTypeConst.cAKBTypeDate
                Case "CHAR", "VARCHAR2", "LONG"
                    C164205DataType(iC164205 + 1) = AKBTypeConst.cAKBTypeString
                End Select
            Next iC164205
        RsC164205_EOF = Not RsC164205.Read()

        GoTo Comp_C164206

    Comp_C164206:

        ' FV_CODIGO MOEDA IS NOT NULL
        sCurrComponent = "FV_CODIGO MOEDA IS NOT NULL"
        C164206DataType = 4
        C164206 = RsC164205_EOF
        GoTo Comp_C164207

    Comp_C164207:

        ' FV_CODIGO MOEDA IS NOT NULL = 1
        sCurrComponent = "FV_CODIGO MOEDA IS NOT NULL = 1"
        C164207 = (fn_ConvertValueCompiled(C164206, C164206DataType, AKB_DecimalPoint, False) = 1)
        C164207DataType = AKBTypeConst.cAKBTypeLogical
        If C164207 Then
            GoTo Comp_C164214
        Else
            GoTo Comp_C164209
        End If

    Comp_C164208:

        ' Cont Gerados
        sCurrComponent = "Cont Gerados"
        C164208 = 0
        C164208DataType = 1
        GoTo Comp_C54079

    Comp_C164209:

        ' Cont Gerados + 1
        sCurrComponent = "Cont Gerados + 1"
        C164209 = fn_ConvertValueCompiled(C164208, C164208DataType, AKB_DecimalPoint, True) + 1
        C164209DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C164210

    Comp_C164210:

        ' Atv_Cont Gerados +1
        sCurrComponent = "Atv_Cont Gerados +1"
        C164210DataType = 4
        C164208 = fn_ConvertValueCompiled(C164209 , 1, AKB_DecimalPoint)
        C164210 = True
        GoTo Comp_C54091

    Comp_C164211:

        ' Cont <> 0
        sCurrComponent = "Cont <> 0"
        C164211 = (fn_ConvertValueCompiled(C164208, C164208DataType, AKB_DecimalPoint, False) <> 0)
        C164211DataType = AKBTypeConst.cAKBTypeLogical
        If C164211 Then
            GoTo Comp_C164216
        Else
            GoTo Comp_C54083
        End If

    Comp_C164212:

        ' MSG1
        sCurrComponent = "MSG1"
        Dim row_C164212 As DataRow = msg.NewRow()
        row_C164212(0) = "Existe um ou mais Pré-Pedido que não pode ser gerado, pois não existe Codigo da Moeda  para ele(s)."
        msg.Rows.Add(row_C164212)
        C164212DataType = AKBTypeConst.cAKBTypeNumeric

        GoTo Comp_C54083

    Comp_C164213:

        ' Total
        sCurrComponent = "Total"
        C164213 = 0
        C164213DataType = 1
        GoTo Comp_C164208

    Comp_C164214:

        ' Total +1
        sCurrComponent = "Total +1"
        C164214 = fn_ConvertValueCompiled(C164213, C164213DataType, AKB_DecimalPoint, True) + 1
        C164214DataType = AKBTypeConst.cAKBTypeNumeric
        GoTo Comp_C164215

    Comp_C164215:

        ' Atv_Total + 1
        sCurrComponent = "Atv_Total + 1"
        C164215DataType = 4
        C164213 = fn_ConvertValueCompiled(C164214 , 1, AKB_DecimalPoint)
        C164215 = True
        GoTo Comp_C54081

    Comp_C164216:

        ' Cont Gerados <= Total
        sCurrComponent = "Cont Gerados <= Total"
        C164216 = (fn_ConvertValueCompiled(C164208, C164208DataType, AKB_DecimalPoint, False) < fn_ConvertValueCompiled(C164213, C164213DataType, AKB_DecimalPoint, False))
        C164216DataType = AKBTypeConst.cAKBTypeLogical
        If C164216 Then
            GoTo Comp_C164212
        Else
            GoTo Comp_C54083
        End If

    Exit_R4555:

        Exit Function

    Err_R4555:
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
